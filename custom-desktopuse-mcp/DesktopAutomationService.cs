using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.UIA3;
using FlaUI.Core.Conditions; // Required for specific search conditions
using System.Diagnostics;
using System.Text.Json;

namespace DesktopMcpServer
{
    public class DesktopAutomationService : IDisposable
    {
        private readonly UIA3Automation _automation;
        private FlaUI.Core.Application? _currentApp;

        // Configurable timeouts
        private readonly TimeSpan _launchTimeout = TimeSpan.FromSeconds(20); 
        private readonly TimeSpan _windowSearchTimeout = TimeSpan.FromSeconds(5);
        private readonly TimeSpan _elementSearchTimeout = TimeSpan.FromSeconds(5);

        public DesktopAutomationService()
        {
            _automation = new UIA3Automation();
        }

        // --- CORE TOOLS ---

        public string LaunchApp(string path)
        {
            try
            {
                //removing logic to allow multiple apps to be launched
                //if (_currentApp != null && !_currentApp.HasExited) CloseApp();

                _currentApp = FlaUI.Core.Application.Launch(path);
                
                var handleFound = _currentApp.WaitWhileMainHandleIsMissing(_launchTimeout);
                if (!handleFound) 
                    return $"Launched: {path} (PID: {_currentApp.ProcessId}), but main window did not appear within 20s. Try using 'GetWindowTree' manually.";

                _currentApp.WaitWhileBusy(); 
                Wait.UntilInputIsProcessed();

                return $"Successfully launched: {path} (PID: {_currentApp.ProcessId})";
            }
            catch (Exception ex) { return $"Error launching app: {ex.Message}"; }
        }

        public string CloseApp(string? windowName = null)
        {
            try
            {
                if (string.IsNullOrEmpty(windowName) || windowName.Equals("current", StringComparison.OrdinalIgnoreCase))
                {
                    if (_currentApp == null || _currentApp.HasExited)
                        return "No application is currently attached.";

                    int pid = _currentApp.ProcessId;
                    try { _currentApp.Close(); } catch { _currentApp.Kill(); }
                    _currentApp.Dispose();
                    _currentApp = null;
                    return $"Closed current app (PID: {pid})";
                }

                var processes = Process.GetProcesses();
                int closedCount = 0;
                
                foreach (var p in processes)
                {
                    if (p.Id <= 4) continue;
                    try
                    {
                        bool match = false;
                        if (p.ProcessName.Contains(windowName, StringComparison.OrdinalIgnoreCase)) match = true;
                        else if (!string.IsNullOrEmpty(p.MainWindowTitle) && 
                                 p.MainWindowTitle.Contains(windowName, StringComparison.OrdinalIgnoreCase)) match = true;

                        if (match)
                        {
                            p.CloseMainWindow();
                            p.WaitForExit(1000); 
                            if (!p.HasExited) p.Kill();
                            closedCount++;
                        }
                    }
                    catch { }
                }
                return closedCount > 0 ? $"Closed {closedCount} apps matching '{windowName}'." : $"No apps found matching '{windowName}'.";
            }
            catch (Exception ex) { return $"Error closing app: {ex.Message}"; }
        }

        public string GetWindowTree(string fieldName)
        {
            try
            {
                var window = FindWindow(fieldName);
                if (window == null) return $"Error: No window found matching '{fieldName}'";

                var tree = BuildSimplifiedTree(window, 0, maxDepth: 4);
                return JsonSerializer.Serialize(tree, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (Exception ex) { return $"Error retrieving tree: {ex.Message}"; }
        }

        // --- ACTION TOOLS ---

        public string ClickElement(string fieldName, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}' in window '{windowName ?? "Current"}'";

                // Handle Offscreen Elements
                if (element.IsOffscreen)
                {
                    if (element.Patterns.ScrollItem.TryGetPattern(out var scrollPattern))
                    {
                        scrollPattern.ScrollIntoView();
                    }
                }

                try { element.WaitUntilClickable(TimeSpan.FromSeconds(2)); } catch {}

                // STRATEGY 1: Background (Invoke Pattern)
                if (element.Patterns.Invoke.TryGetPattern(out var invokePattern))
                {
                    invokePattern.Invoke();
                    return $"Successfully invoked: {fieldName} (Background/Pattern)";
                }

                // STRATEGY 2: Foreground (Physical Mouse Click)
                element.SetForeground();
                element.Click();
                return $"Successfully clicked: {fieldName} (Foreground/Mouse)";
            }
            catch (Exception ex) { return $"Error clicking element: {ex.Message}"; }
        }

        public string RightClickElement(string fieldName, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}' in window '{windowName ?? "Current"}'";

                // STEP 1: Handle Offscreen (Defensive)
                // If the element is hidden in a list, scroll it into view first
                if (element.IsOffscreen)
                {
                    if (element.Patterns.ScrollItem.TryGetPattern(out var scrollPattern))
                    {
                        scrollPattern.ScrollIntoView();
                    }
                }

                // STEP 2: Wait for UI (Defensive)
                // Ensure the element is actually ready to receive input
                try { element.WaitUntilClickable(TimeSpan.FromSeconds(2)); } catch {}

                // STEP 3: Foreground & Click (Physical)
                // There is no "RightClickPattern", so we must use physical input simulation.
                // We force focus first to ensure the context menu appears on top.
                element.SetForeground();
                element.RightClick();
                
                // Optional: Wait a tiny bit for the Context Menu to actually fade in
                Wait.UntilInputIsProcessed();
                Thread.Sleep(200);

                return $"Successfully right-clicked: {fieldName}";
            }
            catch (Exception ex) { return $"Error right-clicking: {ex.Message}"; }
        }

        public string WriteText(string fieldName, string value, string? specialKeys = null, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}'";

                // STRATEGY 1: Background (Value Pattern) - PREFERRED
                // This sets text directly in memory. It prevents "Enter" triggers and works in background.
                // We only use this if we don't have special keys (which imply keyboard interaction).
                if (string.IsNullOrEmpty(specialKeys) && !string.IsNullOrEmpty(value))
                {
                    if (element.Patterns.Value.TryGetPattern(out var valuePattern))
                    {
                        if (valuePattern.IsReadOnly.Value) 
                            return $"Error: Field '{fieldName}' is read-only.";

                        valuePattern.SetValue(value);
                        return $"Successfully set text to '{value}' (Background/Pattern).";
                    }
                }

                // STRATEGY 2: Foreground (Input Simulation) - FALLBACK
                element.SetForeground();
                element.Focus();
                Wait.UntilInputIsProcessed(); 

                // Send Special Keys (using helper)
                if (!string.IsNullOrEmpty(specialKeys))
                {
                    string netKeys = TranslateKeys(specialKeys);
                    System.Windows.Forms.SendKeys.SendWait(netKeys);
                    Wait.UntilInputIsProcessed();
                }

                // Type the value (if Strategy 1 failed)
                if (!string.IsNullOrEmpty(value))
                {
                    // Try ValuePattern again (sometimes focus helps)
                     if (element.Patterns.Value.TryGetPattern(out var valuePattern) && !valuePattern.IsReadOnly.Value)
                    {
                        valuePattern.SetValue(value);
                    }
                    else
                    {
                        // Fallback: Human-like typing
                        // Note: If value contains '~', SendKeys treats it as Enter. Keyboard.Type treats it literal.
                        Keyboard.Type(value);
                    }
                }

                return $"Successfully wrote to '{fieldName}' (Foreground/Focus).";
            }
            catch (Exception ex) { return $"Error writing text: {ex.Message}"; }
        }

        public string GetText(string fieldName, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}'";

                if (element.Patterns.Value.TryGetPattern(out var valuePattern))
                    return valuePattern.Value.Value;
                if (element.Patterns.Text.TryGetPattern(out var textPattern))
                    return textPattern.DocumentRange.GetText(Int32.MaxValue);
                return element.Name;
            }
            catch (Exception ex) { return $"Error reading text: {ex.Message}"; }
        }

        // 1. The Helper Method (Translates "CTRL+V" -> "^v")
        private string TranslateKeys(string keys)
        {
            if (string.IsNullOrEmpty(keys)) return "";

            string input = keys.ToUpper().Trim();

            // Handle Modifiers
            input = input.Replace("CTRL+", "^");
            input = input.Replace("ALT+", "%");
            input = input.Replace("SHIFT+", "+");
            
            // Handle Keywords
            var specialMap = new Dictionary<string, string>
            {
                { "ENTER", "{ENTER}" },
                { "TAB", "{TAB}" },
                { "BS", "{BACKSPACE}" },
                { "BACKSPACE", "{BACKSPACE}" },
                { "DEL", "{DELETE}" },
                { "DELETE", "{DELETE}" },
                { "ESC", "{ESC}" },
                { "DOWN", "{DOWN}" },
                { "UP", "{UP}" },
                { "LEFT", "{LEFT}" },
                { "RIGHT", "{RIGHT}" },
                { "HOME", "{HOME}" },
                { "END", "{END}" },
                { "PGUP", "{PGUP}" },
                { "PGDN", "{PGDN}" },
                { "F5", "{F5}" },
                { "F6", "{F6}" },
                { "F7", "{F7}" },
                { "F8", "{F8}" },
                { "F9", "{F9}" },
                { "F10", "{F10}" },
                { "F11", "{F11}" },
                { "F12", "{F12}" },
            };

            if (specialMap.ContainsKey(input)) return specialMap[input];
            // 4. Handle Case: CTRL +C -> "^C" is valid, but lower case "c" is safer for some apps
            return input;
        }

        // 2. The Tool (Uses the Helper + Focuses Window)
        public string SendSpecialKeys(string specialKeys, string? windowName = null)
        {
            try
            {
                // Translate: "CTRL+V" becomes "^V"
                string netKeys = TranslateKeys(specialKeys);

                // Focus logic
                if (!string.IsNullOrEmpty(windowName))
                {
                    var window = FindWindow(windowName);
                    if (window != null)
                    {
                        window.SetForeground();
                        window.Focus();
                        Wait.UntilInputIsProcessed();
                        Thread.Sleep(200); 
                    }
                    else
                    {
                        return $"Error: Target window '{windowName}' not found.";
                    }
                }

                // Send the translated keys
                System.Windows.Forms.SendKeys.SendWait(netKeys);
                Wait.UntilInputIsProcessed();
                return $"Sent keys: {netKeys} (Original: {specialKeys})";
            }
            catch (Exception ex) { return $"Error sending keys: {ex.Message}"; }
        }

        public string SendKeys(string fieldName, string value, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}'";

                element.SetForeground();
                element.Focus();
                Wait.UntilInputIsProcessed();

                Random rnd = new Random();
                foreach (char c in value)
                {
                    Keyboard.Type(c.ToString());
                    Thread.Sleep(rnd.Next(50, 150)); 
                }
                return $"Typed '{value}' (Human-like).";
            }
            catch (Exception ex) { return $"Error typing: {ex.Message}"; }
        }

public string SelectItems(string fieldName, string value, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}'";

                var itemNames = value.Split(',').Select(i => i.Trim()).ToList();
                var successLog = new List<string>();

                // STRATEGY 1: Value Pattern (Background)
                // Modern combos allow setting the value directly (e.g. "Windows Authentication") 
                // without expanding the dropdown. This is instant and invisible.
                if (itemNames.Count == 1)
                {
                    if (element.Patterns.Value.TryGetPattern(out var valuePattern) && !valuePattern.IsReadOnly.Value)
                    {
                        try 
                        {
                            valuePattern.SetValue(itemNames[0]);
                            // Verification check
                            if (valuePattern.Value.Value == itemNames[0])
                                return $"Selected (via Value Pattern): {itemNames[0]}";
                        }
                        catch {}
                    }
                }

                // STRATEGY 2: Expand & Select (Background-ish)
                // Does not steal focus, but works with UI logic.
                if (element.Patterns.ExpandCollapse.TryGetPattern(out var expand))
                {
                    if (expand.ExpandCollapseState.Value == ExpandCollapseState.Collapsed)
                    {
                        expand.Expand();
                        Thread.Sleep(500); // Wait for render
                        Wait.UntilInputIsProcessed();
                    }
                }

                foreach (var name in itemNames)
                {
                    var child = SmartFindElementInTree(element, name);
                    
                    if (child != null)
                    {
                        // Scroll if needed (Background)
                        if (child.IsOffscreen && child.Patterns.ScrollItem.TryGetPattern(out var scroll))
                            scroll.ScrollIntoView();

                        // Try Selection Pattern (Background)
                        if (child.Patterns.SelectionItem.TryGetPattern(out var selItem))
                        {
                            if (itemNames.Count > 1) selItem.AddToSelection();
                            else selItem.Select();
                            successLog.Add(name);
                        }
                        else
                        {
                            // Fallback: Physical Click (Foreground)
                            child.SetForeground();
                            child.Click();
                            successLog.Add(name);
                        }
                        Thread.Sleep(200); 
                    }
                    else
                    {
                         return $"Error: Item '{name}' not found in list.";
                    }
                }
                return $"Selected: {string.Join(", ", successLog)}";
            }
            catch (Exception ex) { return $"Error selecting items: {ex.Message}"; }
        }

        public string SetCheckbox(string fieldName, string value, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}'";

                bool wantChecked = value.ToLower() is "on" or "true" or "checked" or "yes";
                
                // STRATEGY 1: CheckBox Wrapper (Background)
                var checkbox = element.AsCheckBox();
                if (checkbox != null)
                {
                    if (wantChecked && checkbox.IsChecked != true) checkbox.IsChecked = true;
                    else if (!wantChecked && checkbox.IsChecked != false) checkbox.IsChecked = false;
                    return $"Checkbox set to {(wantChecked ? "Checked" : "Unchecked")} (Pattern)";
                }
                
                // STRATEGY 2: Toggle Pattern (Background)
                if (element.Patterns.Toggle.TryGetPattern(out var toggle))
                {
                    if (wantChecked && toggle.ToggleState.Value != ToggleState.On) toggle.Toggle();
                    else if (!wantChecked && toggle.ToggleState.Value == ToggleState.On) toggle.Toggle();
                    return $"Toggled '{fieldName}' (Pattern)";
                }

                // STRATEGY 3: Click (Foreground Fallback)
                element.SetForeground();
                element.Click();
                return "Toggled via Click (Fallback)";
            }
            catch (Exception ex) { return $"Error setting checkbox: {ex.Message}"; }
        }

        public string SelectRadioButton(string fieldName, string value, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                
                // Logic to find child in group
                if (element != null && !string.IsNullOrEmpty(value) && !(value.ToLower() is "on" or "true"))
                {
                     var childOption = SmartFindElementInTree(element, value);
                     if (childOption != null) element = childOption;
                }

                if (element == null) return $"Error: Radio element not found '{fieldName}'";

                // STRATEGY 1: Selection Pattern (Background)
                if (element.Patterns.SelectionItem.TryGetPattern(out var selectionPattern))
                {
                    selectionPattern.Select();
                    return $"Selected radio: {element.Name ?? fieldName} (Pattern)";
                }
                
                // STRATEGY 2: Click (Foreground)
                element.SetForeground();
                element.Click(); 
                return $"Clicked radio: {element.Name ?? fieldName} (Mouse)";
            }
            catch (Exception ex) { return $"Error selecting radio: {ex.Message}"; }
        }

        public string WaitForElement(string? fieldName, string? windowName = null, int timeoutSeconds = 20)
        {
            try
            {
                var waitTime = TimeSpan.FromSeconds(timeoutSeconds);

                // Scenario 1: Wait for Window Only
                if (string.IsNullOrEmpty(fieldName))
                {
                    if (string.IsNullOrEmpty(windowName)) return "Error: Must provide windowName.";
                    var window = Retry.WhileNull(() => FindWindow(windowName), waitTime).Result;
                    return window != null 
                        ? $"Found window '{windowName}'." 
                        : $"Timeout: Window '{windowName}' not found.";
                }

                // Scenario 2: Wait for Element
                var element = Retry.WhileNull(() =>
                {
                    var root = FindWindow(windowName);
                    if (root == null) return null;
                    return SmartFindElementInTree(root, fieldName);
                }, waitTime).Result;

                return element != null 
                    ? $"Found element '{fieldName}'." 
                    : $"Timeout: Element '{fieldName}' not found.";
            }
            catch (Exception ex) { return $"Error waiting: {ex.Message}"; }
        }

        // --- INTELLIGENT SEARCH HELPERS ---

        private AutomationElement? FindWindow(string? windowName)
        {
            if (string.IsNullOrEmpty(windowName) || windowName.Equals("current", StringComparison.OrdinalIgnoreCase))
            {
                if (_currentApp != null && !_currentApp.HasExited)
                {
                    try { return _currentApp.GetMainWindow(_automation); } catch {}
                }
                return _automation.GetDesktop();
            }

            var desktop = _automation.GetDesktop();
            
            return Retry.WhileNull(() => 
            {
                // 1. Fuzzy Search (Top-Level Only - Fast)
                var topLevelWindows = desktop.FindAllChildren();
                foreach (var win in topLevelWindows)
                {
                    var title = GetSafeProperty(() => win.Name);
                    if (!string.IsNullOrEmpty(title) && title.Contains(windowName, StringComparison.OrdinalIgnoreCase))
                        return win;

                    var id = GetSafeProperty(() => win.AutomationId);
                    if (!string.IsNullOrEmpty(id) && id.Contains(windowName, StringComparison.OrdinalIgnoreCase))
                        return win;

                    try 
                    {
                        var pid = win.Properties.ProcessId.Value;
                        var proc = Process.GetProcessById(pid);
                        if (proc.ProcessName.Contains(windowName, StringComparison.OrdinalIgnoreCase))
                            return win;
                    }
                    catch { }
                }

                // 2. Exact Deep Search (Fallback - Slow but Reliable for Child Dialogs)
                // This fixes "Connect to Server" not being found because it is a child of the main window
                var deepMatch = desktop.FindFirstDescendant(cf => cf.ByName(windowName));
                if (deepMatch != null) return deepMatch;

                // 3. Control Type Fallback
                if (Enum.TryParse<ControlType>(windowName, true, out var type))
                {
                     return desktop.FindFirstDescendant(cf => cf.ByControlType(type));
                }

                return null;
            }, _windowSearchTimeout).Result;
        }

        private AutomationElement? SmartFindElement(string fieldName, string? windowName = null)
        {
            var root = FindWindow(windowName);
            if (root == null) return null;

            Wait.UntilInputIsProcessed(); 
            return Retry.WhileNull(() => SmartFindElementInTree(root, fieldName), _elementSearchTimeout).Result;
        }

        private AutomationElement? SmartFindElementInTree(AutomationElement root, string fieldName)
        {
            var cf = _automation.ConditionFactory;

            // 1. Exact ID
            var el = root.FindFirstDescendant(cf.ByAutomationId(fieldName));
            if (el != null) return el;

            // 2. Exact Name
            el = root.FindFirstDescendant(cf.ByName(fieldName));
            if (el != null) return el;

            // 3. Case Insensitive Name (Using PropertyCondition directly for fuzzy match)
            try {
                var nameCond = new PropertyCondition(_automation.PropertyLibrary.Element.Name, fieldName, PropertyConditionFlags.IgnoreCase);
                el = root.FindFirstDescendant(nameCond);
                if (el != null) return el;
            } catch {}

            // 4. Partial Match (SAFE)
            // Only scans logical elements to avoid timeouts. Fixes "_New Query" vs "New Query"
            try 
            {
                var typeCond = cf.ByControlType(ControlType.Button)
                    .Or(cf.ByControlType(ControlType.Edit))
                    .Or(cf.ByControlType(ControlType.Text))
                    .Or(cf.ByControlType(ControlType.Document))
                    .Or(cf.ByControlType(ControlType.MenuItem))
                    .Or(cf.ByControlType(ControlType.TreeItem))
                    .Or(cf.ByControlType(ControlType.ListItem))
                    .Or(cf.ByControlType(ControlType.CheckBox));
                
                var candidates = root.FindAllDescendants(typeCond);
                el = candidates.FirstOrDefault(e => e.Name != null && e.Name.Contains(fieldName, StringComparison.OrdinalIgnoreCase));
                if (el != null) return el;
            }
            catch {}

            // 5. Class/Role
            el = root.FindFirstDescendant(cf.ByClassName(fieldName));
            if (el != null) return el;

            if (Enum.TryParse<ControlType>(fieldName, true, out var type))
            {
                el = root.FindFirstDescendant(cf.ByControlType(type));
                if (el != null) return el;
            }

            return null;
        }
        
        private SimplifiedControl BuildSimplifiedTree(AutomationElement element, int currentDepth, int maxDepth)
        {
            var node = new SimplifiedControl
            {
                ControlType = GetSafeProperty(() => element.ControlType.ToString()),
                Name = GetSafeProperty(() => element.Name),
                AutomationId = GetSafeProperty(() => element.AutomationId),
                ClassName = GetSafeProperty(() => element.ClassName)
            };

            if (currentDepth < maxDepth)
            {
                try
                {
                    var children = element.FindAllChildren();
                    foreach (var child in children)
                    {
                        var cType = GetSafeProperty(() => child.ControlType.ToString());
                        var cName = GetSafeProperty(() => child.Name);
                        var cId = GetSafeProperty(() => child.AutomationId);

                        if (cType == ControlType.Pane.ToString() && string.IsNullOrEmpty(cName) && string.IsNullOrEmpty(cId))
                            continue;
                        node.Children.Add(BuildSimplifiedTree(child, currentDepth + 1, maxDepth));
                    }
                }
                catch { }
            }
            return node;
        }

        private string GetSafeProperty(Func<object> propertyGetter)
        {
            try { return propertyGetter()?.ToString() ?? ""; }
            catch { return ""; }
        }

        public void Dispose()
        {
            _currentApp?.Dispose();
            _automation?.Dispose();
        }
    }

    public class SimplifiedControl
    {
        public string ControlType { get; set; } = "";
        public string Name { get; set; } = "";
        public string AutomationId { get; set; } = "";
        public string ClassName { get; set; } = "";
        public List<SimplifiedControl> Children { get; set; } = new();
    }
}
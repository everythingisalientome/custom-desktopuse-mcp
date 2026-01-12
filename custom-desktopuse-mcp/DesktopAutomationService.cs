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
                if (_currentApp != null && !_currentApp.HasExited) CloseApp();

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

                //Adding logic to improve reliability of clicks
                // FIX 1: Handle Offscreen Elements (e.g. at bottom of a list)
                if (element.IsOffscreen)
                {
                    if (element.Patterns.ScrollItem.TryGetPattern(out var scrollPattern))
                    {
                        scrollPattern.ScrollIntoView();
                    }
                }

                // FIX 2: Wait for Animation (e.g. Combobox sliding down)
                // This prevents clicking "air" while the UI is still fading in
                try 
                { 
                    element.WaitUntilClickable(TimeSpan.FromSeconds(2)); 
                } 
                catch {}

                // FIX 3: Ensure Focus
                element.SetForeground();

                if (element.Patterns.Invoke.TryGetPattern(out var invokePattern))
                {
                    invokePattern.Invoke();
                    return $"Successfully invoked: {fieldName}";
                }
                element.Click();
                return $"Successfully clicked: {fieldName}";
            }
            catch (Exception ex) { return $"Error clicking: {ex.Message}"; }
        }

        public string RightClickElement(string fieldName, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}'";
                element.RightClick();
                return $"Right-clicked: {fieldName}";
            }
            catch (Exception ex) { return $"Error right-clicking: {ex.Message}"; }
        }

        public string WriteText(string fieldName, string value, string? specialKeys = null, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}'";

                element.SetForeground();
                element.Focus();
                Wait.UntilInputIsProcessed(); 

                if (!string.IsNullOrEmpty(specialKeys))
                {
                    Keyboard.Type(specialKeys);
                    Wait.UntilInputIsProcessed();
                }

                if (!string.IsNullOrEmpty(value))
                {
                    if (element.Patterns.Value.TryGetPattern(out var valuePattern))
                        valuePattern.SetValue(value);
                    else
                        Keyboard.Type(value);
                }
                return $"Wrote to '{fieldName}'.";
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

        public string SendSpecialKeys(string specialKeys, string? windowName = null)
        {
            try
            {
                //Keyboard.Type(specialKeys);
                // FIX: If windowName is provided, find and focus it FIRST
                if (!string.IsNullOrEmpty(windowName))
                {
                    var window = FindWindow(windowName);
                    if (window != null)
                    {
                        window.SetForeground();
                        window.Focus();
                        Wait.UntilInputIsProcessed();
                        Thread.Sleep(200); // Small pause for OS focus switch
                    }
                    else
                    {
                        return $"Error: Target window '{windowName}' not found for sending keys.";
                    }
                }

                // Use Windows Forms SendKeys to support codes like "{ENTER}", "^a"
                System.Windows.Forms.SendKeys.SendWait(specialKeys);
                Wait.UntilInputIsProcessed();
                return $"Sent keys: {specialKeys}";
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

                var comboBox = element.AsComboBox();
                if (comboBox != null && itemNames.Count == 1)
                {
                    //comboBox.Select(itemNames[0]);
                    //return $"Selected: {itemNames[0]}";
                    // FIX: Ensure we wait after expanding, or use the pattern directly if collapsed
                    if (comboBox.Patterns.ExpandCollapse.TryGetPattern(out var expandPattern) &&
                        expandPattern.ExpandCollapseState.Value == ExpandCollapseState.Collapsed)
                    {
                        expandPattern.Expand();
                        Thread.Sleep(500); // FIX: Wait for animation to finish opening
                    }
                    
                    var item = comboBox.Select(itemNames[0]);
                    return $"Selected dropdown item: {itemNames[0]}";
                }

                if (element.Patterns.ExpandCollapse.TryGetPattern(out var expand))
                {
                    if (expand.ExpandCollapseState.Value == ExpandCollapseState.Collapsed)
                    {
                        expand.Expand();
                        // FIX: Explicit pause to allow the dropdown list to render visually
                        Thread.Sleep(500);
                        Wait.UntilInputIsProcessed();
                    }
                }

                foreach (var name in itemNames)
                {
                    var child = SmartFindElementInTree(element, name);
                    //if (child != null && child.Patterns.SelectionItem.TryGetPattern(out var selItem))
                    //{
                    //    if (itemNames.Count > 1) selItem.AddToSelection();
                    //    else selItem.Select();
                    //    successLog.Add(name);
                    //}
                    if (child != null)
                    {
                        // Try selection pattern first (more reliable than clicking)
                        if (child.Patterns.SelectionItem.TryGetPattern(out var selItem))
                        {
                            if (itemNames.Count > 1) selItem.AddToSelection();
                            else selItem.Select();
                            successLog.Add(name);
                        }
                        else
                        {
                            // Fallback to click, but verify visibility
                            child.SetForeground();
                            child.Click();
                            successLog.Add(name);
                        }
                        // Small pause between multiple selections
                        Thread.Sleep(200); 
                    }
                    else
                    {
                         // If not found, it might be off-screen (virtualized). 
                         // Try sending a "Down" key to scroll, then look again? 
                         // For now, just log failure.
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
                
                var checkbox = element.AsCheckBox();
                if (checkbox != null)
                {
                    if (wantChecked && checkbox.IsChecked != true) checkbox.IsChecked = true;
                    else if (!wantChecked && checkbox.IsChecked != false) checkbox.IsChecked = false;
                    return $"Checkbox set to {(wantChecked ? "Checked" : "Unchecked")}";
                }
                
                if (element.Patterns.Toggle.TryGetPattern(out var toggle))
                {
                    if (wantChecked && toggle.ToggleState.Value != ToggleState.On) toggle.Toggle();
                    else if (!wantChecked && toggle.ToggleState.Value == ToggleState.On) toggle.Toggle();
                    return $"Toggled '{fieldName}'";
                }
                return "Error: Element is not a checkbox/toggle.";
            }
            catch (Exception ex) { return $"Error setting checkbox: {ex.Message}"; }
        }

        public string SelectRadioButton(string fieldName, string value, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                
                if (element != null && !string.IsNullOrEmpty(value) && !(value.ToLower() is "on" or "true"))
                {
                     var childOption = SmartFindElementInTree(element, value);
                     if (childOption != null) element = childOption;
                }

                if (element == null) return $"Error: Radio element not found '{fieldName}'";

                if (element.Patterns.SelectionItem.TryGetPattern(out var selectionPattern))
                {
                    selectionPattern.Select();
                    return $"Selected radio: {element.Name ?? fieldName}";
                }
                
                element.Click(); 
                return $"Clicked radio: {element.Name ?? fieldName}";
            }
            catch (Exception ex) { return $"Error selecting radio: {ex.Message}"; }
        }

        public string WaitForElement(string fieldName, string? windowName = null, int timeoutSeconds = 20)
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
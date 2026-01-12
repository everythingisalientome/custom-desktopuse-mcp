using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.UIA3;
using System.Diagnostics;
using System.Text.Json;

namespace DesktopMcpServer
{
    public class DesktopAutomationService : IDisposable
    {
        private readonly UIA3Automation _automation;
        private FlaUI.Core.Application? _currentApp;

        // Configurable timeouts
        private readonly TimeSpan _launchTimeout = TimeSpan.FromSeconds(20); // INCREASED from 2s
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
                if (_currentApp != null && !_currentApp.HasExited)
                {
                    CloseApp();
                }

                _currentApp = FlaUI.Core.Application.Launch(path);
                
                // FIX 1: Increased wait time for main window (20s)
                // This covers splash screens that take a while to close/swap
                var handleFound = _currentApp.WaitWhileMainHandleIsMissing(_launchTimeout);
                
                if (!handleFound) 
                {
                    // If we timed out, we don't crash, but we warn the agent.
                    // Often the app is actually open (e.g., in tray), just no "Main Window" yet.
                    return $"Launched: {path} (PID: {_currentApp.ProcessId}), but main window did not appear within 20 seconds. You may need to use 'GetWindowTree' to find it manually.";
                }

                // FIX 2: Wait for application to be "Idle" (finished loading)
                _currentApp.WaitWhileBusy(); 
                Wait.UntilInputIsProcessed();

                return $"Successfully launched: {path} (PID: {_currentApp.ProcessId})";
            }
            catch (Exception ex)
            {
                return $"Error launching app: {ex.Message}";
            }
        }

        // --- STRATEGY C: VERIFIED LAUNCH ---
        // Use this if standard LaunchApp fails due to Splash Screens
        public string LaunchAppVerified(string path, string expectedTitle)
        {
            try
            {
                if (_currentApp != null && !_currentApp.HasExited)
                    CloseApp();

                _currentApp = FlaUI.Core.Application.Launch(path);
                
                // Custom Retry Loop: Wait up to 30 seconds for specific window title
                var desktop = _automation.GetDesktop();
                var mainWindow = Retry.WhileNull(() =>
                {
                    // Search the entire desktop for a window containing the expected title
                    var found = desktop.FindFirstDescendant(cf => cf.ByName(expectedTitle))?.AsWindow();
                    
                    // Or try by AutomationID if Name failed
                    if (found == null)
                        found = desktop.FindFirstDescendant(cf => cf.ByAutomationId(expectedTitle))?.AsWindow();

                    return found;
                }, TimeSpan.FromSeconds(30)).Result;

                if (mainWindow == null)
                {
                    return $"Launched app (PID: {_currentApp.ProcessId}) but could not find window '{expectedTitle}' after 30 seconds.";
                }

                // Force focus to ensure it's ready
                mainWindow.SetForeground();
                Wait.UntilInputIsProcessed();

                return $"Successfully launched and verified window: '{expectedTitle}'";
            }
            catch (Exception ex)
            {
                return $"Error in verified launch: {ex.Message}";
            }
        }

        // Tool: CloseApp (Unified)
        public string CloseApp(string? windowName = null)
        {
            try
            {
                // Scenario 1: Close the "Current" attached app (Default)
                if (string.IsNullOrEmpty(windowName) || windowName.Equals("current", StringComparison.OrdinalIgnoreCase))
                {
                    if (_currentApp == null || _currentApp.HasExited)
                        return "No application is currently attached to the session.";

                    int pid = _currentApp.ProcessId;
                    try { _currentApp.Close(); } catch { _currentApp.Kill(); }
                    
                    _currentApp.Dispose();
                    _currentApp = null;
                    return $"Successfully closed current application (PID: {pid})";
                }

                // Scenario 2: Find apps by Name or Window Title
                var processes = Process.GetProcesses();
                int closedCount = 0;
                var log = new List<string>();

                foreach (var p in processes)
                {
                    // Skip system processes
                    if (p.Id <= 4) continue;

                    try
                    {
                        bool match = false;
                        
                        // Check Process Name (e.g. "notepad")
                        if (p.ProcessName.Contains(windowName, StringComparison.OrdinalIgnoreCase)) match = true;
                        
                        // Check Window Title (e.g. "Untitled - Notepad")
                        else if (!string.IsNullOrEmpty(p.MainWindowTitle) && 
                                p.MainWindowTitle.Contains(windowName, StringComparison.OrdinalIgnoreCase)) match = true;

                        if (match)
                        {
                            p.CloseMainWindow();
                            p.WaitForExit(1000); 
                            if (!p.HasExited) p.Kill();

                            closedCount++;
                            log.Add($"{p.ProcessName} (PID: {p.Id})");
                        }
                    }
                    catch { }
                }

                if (closedCount == 0)
                    return $"Could not find any running application matching '{windowName}'.";

                return $"Successfully closed {closedCount} application(s): {string.Join(", ", log)}";
            }
            catch (Exception ex)
            {
                return $"Error closing app: {ex.Message}";
            }
        }

        public string GetWindowTree(string fieldName)
        {
            try
            {
                // In this context, 'fieldName' is the criteria to find the window
                var window = FindWindow(fieldName);

                if (window == null) return $"Error: No window found matching '{fieldName}'";

                var tree = BuildSimplifiedTree(window, 0, maxDepth: 4);
                return JsonSerializer.Serialize(tree, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (Exception ex)
            {
                return $"Error retrieving window tree: {ex.Message}";
            }
        }

        // Tool: WaitForElement (NEW)
        public string WaitForElement(string fieldName, string? windowName = null, int timeoutSeconds = 20)
        {
            try
            {
                var waitTime = TimeSpan.FromSeconds(timeoutSeconds);

                // Scenario 1: Wait for Window Only (if fieldName is empty/null)
                if (string.IsNullOrEmpty(fieldName))
                {
                    if (string.IsNullOrEmpty(windowName)) return "Error: Must provide at least fieldName or windowName.";
                    
                    var window = Retry.WhileNull(() => FindWindow(windowName), waitTime).Result;
                    
                    return window != null 
                        ? $"Successfully found window '{windowName}' within {timeoutSeconds}s." 
                        : $"Timeout: Window '{windowName}' did not appear after {timeoutSeconds}s.";
                }

                // Scenario 2: Wait for Element (inside specific Window if provided)
                // We use a custom retry loop here to respect the requested 'timeoutSeconds'
                var element = Retry.WhileNull(() =>
                {
                    // 1. Find/Wait for Window first (fast check)
                    var root = FindWindow(windowName);
                    if (root == null) return null;

                    // 2. Look for the element
                    return SmartFindElementInTree(root, fieldName);
                }, waitTime).Result;

                return element != null 
                    ? $"Successfully found element '{fieldName}' within {timeoutSeconds}s." 
                    : $"Timeout: Element '{fieldName}' did not appear after {timeoutSeconds}s.";
            }
            catch (Exception ex)
            {
                return $"Error waiting for element: {ex.Message}";
            }
        }

        // --- ACTION TOOLS ---

        public string ClickElement(string fieldName, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}' in window '{windowName ?? "Current"}'";

                if (element.Patterns.Invoke.TryGetPattern(out var invokePattern))
                {
                    invokePattern.Invoke();
                    return $"Successfully invoked: {fieldName}";
                }

                element.Click();
                return $"Successfully clicked: {fieldName}";
            }
            catch (Exception ex) { return $"Error clicking element: {ex.Message}"; }
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

        // --- INPUT TOOLS ---

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
                    {
                        valuePattern.SetValue(value);
                    }
                    else
                    {
                        Keyboard.Type(value);
                    }
                }

                return $"Successfully wrote to '{fieldName}'.";
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

        public string SendSpecialKeys(string specialKeys)
        {
            try
            {
                Keyboard.Type(specialKeys);
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

                return $"Successfully typed '{value}' into '{fieldName}' (Human-like).";
            }
            catch (Exception ex) { return $"Error typing human-like text: {ex.Message}"; }
        }

        // --- SELECTION TOOLS ---

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
                    comboBox.Select(itemNames[0]);
                    return $"Selected dropdown item: {itemNames[0]}";
                }

                if (element.Patterns.ExpandCollapse.TryGetPattern(out var expand))
                {
                    if (expand.ExpandCollapseState.Value == ExpandCollapseState.Collapsed)
                    {
                        expand.Expand();
                        Wait.UntilInputIsProcessed();
                    }
                }

                foreach (var name in itemNames)
                {
                    var child = SmartFindElementInTree(element, name);
                    if (child != null && child.Patterns.SelectionItem.TryGetPattern(out var selItem))
                    {
                        if (itemNames.Count > 1) selItem.AddToSelection();
                        else selItem.Select();
                        successLog.Add(name);
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
                    return $"Checkbox '{fieldName}' set to {(wantChecked ? "Checked" : "Unchecked")}";
                }
                
                if (element.Patterns.Toggle.TryGetPattern(out var toggle))
                {
                    if (wantChecked && toggle.ToggleState.Value != ToggleState.On) toggle.Toggle();
                    else if (!wantChecked && toggle.ToggleState.Value == ToggleState.On) toggle.Toggle();
                    return $"Toggled '{fieldName}'";
                }

                return "Error: Element does not support Checkbox/Toggle pattern";
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

        // --- INTELLIGENT SEARCH HELPERS ---

        private AutomationElement? FindWindow(string? windowName)
        {
            if (string.IsNullOrEmpty(windowName) || windowName.Equals("current", StringComparison.OrdinalIgnoreCase))
            {
                if (_currentApp != null && !_currentApp.HasExited)
                {
                    return _currentApp.GetMainWindow(_automation);
                }
                return _automation.GetDesktop();
            }

            var desktop = _automation.GetDesktop();
            return Retry.WhileNull(() => 
            {
                var w = desktop.FindFirstDescendant(cf => cf.ByName(windowName));
                if (w != null) return w;

                w = desktop.FindFirstDescendant(cf => cf.ByAutomationId(windowName));
                if (w != null) return w;
                
                w = desktop.FindFirstDescendant(cf => cf.ByClassName(windowName));
                if (w != null) return w;

                if (Enum.TryParse<ControlType>(windowName, true, out var type))
                {
                    w = desktop.FindFirstDescendant(cf => cf.ByControlType(type));
                    if (w != null) return w;
                }

                return null;
            }, _windowSearchTimeout).Result;
        }

        private AutomationElement? SmartFindElement(string fieldName, string? windowName = null)
        {
            var root = FindWindow(windowName);
            if (root == null) return null;

            Wait.UntilInputIsProcessed(); 

            return Retry.WhileNull(() => SmartFindElementInTree(root, fieldName), 
                                   _elementSearchTimeout).Result;
        }

        private AutomationElement? SmartFindElementInTree(AutomationElement root, string fieldName)
        {
            var el = root.FindFirstDescendant(cf => cf.ByAutomationId(fieldName));
            if (el != null) return el;

            el = root.FindFirstDescendant(cf => cf.ByName(fieldName));
            if (el != null) return el;

            el = root.FindFirstDescendant(cf => cf.ByClassName(fieldName));
            if (el != null) return el;

            if (Enum.TryParse<ControlType>(fieldName, true, out var type))
            {
                el = root.FindFirstDescendant(cf => cf.ByControlType(type));
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
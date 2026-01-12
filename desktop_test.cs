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
                // Wait for the main handle to be created (Initial "Screen Refresh" wait)
                _currentApp.WaitWhileMainHandleIsMissing(TimeSpan.FromSeconds(2));
                return $"Successfully launched: {path} (PID: {_currentApp.ProcessId})";
            }
            catch (Exception ex)
            {
                return $"Error launching app: {ex.Message}";
            }
        }

        public string CloseApp()
        {
            try
            {
                if (_currentApp == null || _currentApp.HasExited)
                    return "No application is currently attached.";

                int pid = _currentApp.ProcessId;
                try { _currentApp.Close(); } catch { _currentApp.Kill(); }
                
                _currentApp.Dispose();
                _currentApp = null;
                
                return $"Successfully closed application (PID: {pid})";
            }
            catch (Exception ex) { return $"Error closing app: {ex.Message}"; }
        }

        public string GetWindowTree(string fieldName)
        {
            try
            {
                // Logic: fieldName here IS the window name/criteria
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

        // --- ACTION TOOLS (Now accepting windowName) ---

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

        public string WriteText(string fieldName, string value, string? specialKeys = null, string? windowName = null)
        {
            try
            {
                var element = SmartFindElement(fieldName, windowName);
                if (element == null) return $"Error: Element not found '{fieldName}'";

                element.SetForeground();
                element.Focus();
                Wait.UntilInputIsProcessed(); // Wait for focus to settle

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
                // Sends to whatever is globally focused (no window search needed)
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
                        Wait.UntilInputIsProcessed(); // Wait for dropdown to actually open
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
                
                // Group logic
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

        // NEW: Helper to find specific window (or default)
        private AutomationElement? FindWindow(string? windowName)
        {
            // 1. If no name provided, use Current App's Main Window
            if (string.IsNullOrEmpty(windowName) || windowName.Equals("current", StringComparison.OrdinalIgnoreCase))
            {
                if (_currentApp != null && !_currentApp.HasExited)
                {
                    return _currentApp.GetMainWindow(_automation);
                }
                // If no app attached, return Desktop to search everything
                return _automation.GetDesktop();
            }

            // 2. If name provided, Search Desktop for that Window
            var desktop = _automation.GetDesktop();
            return Retry.WhileNull(() => 
            {
                // Try Name, AutomationId, ClassName
                return desktop.FindFirstDescendant(cf => cf.ByName(windowName)) 
                    ?? desktop.FindFirstDescendant(cf => cf.ByAutomationId(windowName))
                    ?? desktop.FindFirstDescendant(cf => cf.ByClassName(windowName));
            }, _windowSearchTimeout).Result;
        }

        private AutomationElement? SmartFindElement(string fieldName, string? windowName = null)
        {
            // 1. Resolve Root Window
            var root = FindWindow(windowName);
            if (root == null) return null; // Window itself not found

            // 2. Wait for Screen Refresh / Stability
            // This waits for the application's message queue to be empty (idle)
            Wait.UntilInputIsProcessed(); 

            // 3. Retry Loop (The "Wait and Try Again" logic)
            // Tries repeatedly for 5 seconds to find the element
            return Retry.WhileNull(() => SmartFindElementInTree(root, fieldName), 
                                   _elementSearchTimeout).Result;
        }

        private AutomationElement? SmartFindElementInTree(AutomationElement root, string fieldName)
        {
            // 1. Exact ID
            var el = root.FindFirstDescendant(cf => cf.ByAutomationId(fieldName));
            if (el != null) return el;

            // 2. Exact Name
            el = root.FindFirstDescendant(cf => cf.ByName(fieldName));
            if (el != null) return el;

            // 3. Class Name
            el = root.FindFirstDescendant(cf => cf.ByClassName(fieldName));
            if (el != null) return el;

            // 4. Control Type
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
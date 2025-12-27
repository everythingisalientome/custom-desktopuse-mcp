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
        // Store the currently controlled application here
        private FlaUI.Core.Application? _currentApp;
        public DesktopAutomationService()
        {
            _automation = new UIA3Automation();
        }

        
        // Tool: Launch Application
        public string LaunchApp(string path)
        {
            try
            {
                //var process = Process.Start(path);
                //return $"Successfully launched: {path} (PID: {process?.Id})";
                // Dispose previous app instance if exists to prevent memory leaks
                _currentApp?.Dispose();

                // Launch app with FlaUI
                // This gives us a direct handle to the process/application object
                _currentApp = FlaUI.Core.Application.Launch(path);
                
                //Wait a moment for the main window to be created
                _currentApp.WaitWhileMainHandleIsMissing(TimeSpan.FromSeconds(2));

                return $"Successfully launched: {path} (PID: {_currentApp.ProcessId}). Application object cached for future commands.";
            }
            catch (Exception ex)
            {
                return $"Error launching app: {ex.Message}";
            }
        }

        // Tool: Get the window tree as JSON
        public string GetWindowTree(string windowTitle)
        {
            try
            {
                Window? window = null;
                // Logic 1: Use Cached App (Fastest & Most Reliable)
                // If the user asks for "Current", "Active", or the title matches the launched app
                if (_currentApp != null && !_currentApp.HasExited)
                {
                    // Try to get the main window from the cached app object
                    var cachedWindow = _currentApp.GetMainWindow(_automation);
                    
                    // If the user didn't specify a title, OR the titles match, use this one
                    if (string.IsNullOrEmpty(windowTitle) || 
                        windowTitle.Equals("current", StringComparison.OrdinalIgnoreCase) ||
                        (cachedWindow != null && cachedWindow.Name.Contains(windowTitle, StringComparison.OrdinalIgnoreCase)))
                    {
                        window = cachedWindow;
                    }
                }

                // Logic 2: Search Desktop (Fallback)
                // If we didn't find it in cache, search the whole desktop like before
                if (window == null && !string.IsNullOrEmpty(windowTitle))
                {
                    var desktop = _automation.GetDesktop();
                    window = Retry.WhileNull(() =>
                    {
                        var exact = desktop.FindFirstDescendant(cf => cf.ByName(windowTitle))?.AsWindow();
                        if (exact != null) return exact;
                        var allWindows = desktop.FindAllChildren();
                        return allWindows.FirstOrDefault(w => w.Name.Contains(windowTitle, StringComparison.OrdinalIgnoreCase))?.AsWindow();
                    }, TimeSpan.FromSeconds(2)).Result;
                }

                if (window == null) return $"Error: No window found containing title '{windowTitle}'";

                // Build JSON Tree
                var tree = BuildSimplifiedTree(window, 0, maxDepth: 4);
                return JsonSerializer.Serialize(tree, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (Exception ex)
            {
                return $"Error retrieving window tree: {ex.Message}";
            }
        }

        //Pass this as criteria,Pass this as value
        //Field Name / x:Name,"""id""","""createReportButton"""
        //AutomationId,"""id""","""createReportButton"""
        //Visible Text / Label,"""name""","""Create Report"""
        //Button Text,"""name""","""Submit"""

        // Tool: Click Element
        public string ClickElement(string criteria, string value)
        {
            try
            {
                var element = FindElement(criteria, value);
                if (element == null) return $"Error: Element not found ({criteria}={value})";

                if (element.Patterns.Invoke.TryGetPattern(out var invokePattern))
                {
                    invokePattern.Invoke();
                    return $"Successfully invoked element: {value}";
                }

                element.Click();
                return $"Successfully clicked element: {value}";
            }
            catch (Exception ex)
            {
                return $"Error clicking element: {ex.Message}";
            }
        }

        // Tool: RightClick
        public string RightClickElement(string criteria, string value)
        {
            try
            {
                var element = FindElement(criteria, value);
                if (element == null) return $"Error: Element not found ({criteria}={value})";

                element.RightClick();
                return $"Successfully right-clicked element: {value}";
            }
            catch (Exception ex)
            {
                return $"Error right-clicking element: {ex.Message}";
            }
        }

        // Tool: GetText
        public string GetText(string criteria, string value)
        {
            try
            {
                var element = FindElement(criteria, value);
                if (element == null) return $"Error: Element not found ({criteria}={value})";

                // 1. Try ValuePattern (TextBoxes)
                if (element.Patterns.Value.TryGetPattern(out var valuePattern))
                {
                    return valuePattern.Value.Value;
                }
                
                // 2. Try TextPattern (Documents/RichText)
                if (element.Patterns.Text.TryGetPattern(out var textPattern))
                {
                    return textPattern.DocumentRange.GetText(Int32.MaxValue);
                }

                // 3. Fallback to Name property (Labels/Buttons often store text here)
                return element.Name;
            }
            catch (Exception ex)
            {
                return $"Error reading text: {ex.Message}";
            }
        }

        // Tool: TypeText (Renamed/Refined from WriteText)
        public string TypeText(string criteria, string value, string text)
        {
            try
            {
                var element = FindElement(criteria, value);
                if (element == null) return $"Error: Element not found ({criteria}={value})";

                // Always try setting value programmatically first (instant)
                if (element.Patterns.Value.TryGetPattern(out var valuePattern))
                {
                    valuePattern.SetValue(text);
                    return $"Successfully set value to: {text}";
                }

                // Fallback to typing
                element.SetForeground();
                element.Focus();
                Wait.UntilInputIsProcessed();
                Keyboard.Type(text);
                return $"Successfully typed text: {text}";
            }
            catch (Exception ex)
            {
                return $"Error typing text: {ex.Message}";
            }
        }

        // Tool: SendKeys (Global keys)
        public string SendKeys(string keys)
        {
            try
            {
                // NOTE: This sends keys to whatever is currently focused!
                Keyboard.Type(keys);
                return $"Successfully sent keys: {keys}";
            }
            catch (Exception ex)
            {
                return $"Error sending keys: {ex.Message}";
            }
        }

        // Tool: SelectItems
        public string SelectItems(string criteria, string value, string items)
        {
            try
            {
                var element = FindElement(criteria, value);
                if (element == null) return $"Error: Element not found ({criteria}={value})";

                var itemNames = items.Split(',').Select(i => i.Trim()).ToList();
                var successLog = new List<string>();

                // Strategy 1: Handle as ComboBox (The standard "Dropdown")
                // FlaUI's ComboBox wrapper handles expanding/collapsing automatically
                var comboBox = element.AsComboBox();
                if (comboBox != null && itemNames.Count == 1)
                {
                    var item = comboBox.Select(itemNames[0]);
                    if (item != null) return $"Successfully selected dropdown item: {itemNames[0]}";
                }

                // Strategy 2: Handle as Generic List/Container (For Multi-select or non-standard dropdowns)
                // If it's a ComboBox that wasn't handled above, ensure it's expanded
                if (element.Patterns.ExpandCollapse.TryGetPattern(out var expandPattern))
                {
                    if (expandPattern.ExpandCollapseState.Value == ExpandCollapseState.Collapsed)
                    {
                        expandPattern.Expand();
                        Wait.UntilInputIsProcessed();
                    }
                }

                foreach (var name in itemNames)
                {
                    // Find the specific item (child) by name
                    var childItem = Retry.WhileNull(() =>
                    {
                        return element.FindFirstDescendant(cf => cf.ByName(name));
                    }, TimeSpan.FromSeconds(1)).Result;

                    if (childItem == null)
                    {
                        successLog.Add($"Failed to find item: '{name}'");
                        continue;
                    }

                    // Try to select it
                    if (childItem.Patterns.SelectionItem.TryGetPattern(out var selectionPattern))
                    {
                        if (itemNames.Count > 1)
                            selectionPattern.AddToSelection(); // Multi-select
                        else
                            selectionPattern.Select(); // Single select
                        
                        successLog.Add($"Selected: {name}");
                    }
                    else
                    {
                        // Fallback: Just click it (sometimes works for simple custom dropdowns)
                        childItem.Click();
                        successLog.Add($"Clicked: {name}");
                    }
                }

                return string.Join("; ", successLog);
            }
            catch (Exception ex)
            {
                return $"Error selecting items: {ex.Message}";
            }
        }

        // Tool: SetCheckbox
        public string SetCheckbox(string criteria, string value, string toggleState)
        {
            try
            {
                var element = FindElement(criteria, value);
                if (element == null) return $"Error: Element not found ({criteria}={value})";

                // Parse desired state ("on", "true", "checked" -> true)
                bool wantChecked = toggleState.ToLower() is "on" or "true" or "checked" or "yes";

                // Try using the specific CheckBox wrapper first (easiest)
                var checkbox = element.AsCheckBox();
                if (checkbox != null)
                {
                    if (wantChecked && checkbox.IsChecked != true)
                    {
                        checkbox.IsChecked = true; // Use FlaUI helper to toggle if needed
                        return $"Successfully checked: {value}";
                    }
                    else if (!wantChecked && checkbox.IsChecked != false)
                    {
                        checkbox.IsChecked = false;
                        return $"Successfully unchecked: {value}";
                    }
                    return $"Checkbox '{value}' was already in the desired state.";
                }

                // Fallback: Raw Toggle Pattern (for custom controls)
                if (element.Patterns.Toggle.TryGetPattern(out var togglePattern))
                {
                    var currentState = togglePattern.ToggleState.Value;
                    
                    if (wantChecked && currentState != ToggleState.On)
                    {
                        togglePattern.Toggle();
                        return $"Successfully toggled ON: {value}";
                    }
                    else if (!wantChecked && currentState == ToggleState.On)
                    {
                        togglePattern.Toggle();
                        return $"Successfully toggled OFF: {value}";
                    }
                    return $"Element '{value}' was already in correct state.";
                }

                return "Error: Element does not support Toggling/Checkbox pattern.";
            }
            catch (Exception ex)
            {
                return $"Error setting checkbox: {ex.Message}";
            }
        }

        // Tool: SelectRadioButton
        public string SelectRadioButton(string criteria, string value)
        {
            try
            {
                var element = FindElement(criteria, value);
                if (element == null) return $"Error: Element not found ({criteria}={value})";

                // 1. Try SelectionItem Pattern (Standard Radio Button logic)
                if (element.Patterns.SelectionItem.TryGetPattern(out var selectionPattern))
                {
                    selectionPattern.Select();
                    return $"Successfully selected radio button: {value}";
                }
                
                // 2. Fallback: Just click it (Works for simple web-style radio buttons)
                element.Click();
                return $"Clicked radio button (fallback): {value}";
            }
            catch (Exception ex)
            {
                return $"Error selecting radio button: {ex.Message}";
            }
        }

        // --- Helper Methods ---

        private AutomationElement? FindElement(string criteria, string value)
        {
            var desktop = _automation.GetDesktop();
            return Retry.WhileNull(() =>
            {
                return criteria.ToLower() switch
                {
                    "id" or "automationid" => desktop.FindFirstDescendant(cf => cf.ByAutomationId(value)),
                    "name" => desktop.FindFirstDescendant(cf => cf.ByName(value)),
                    _ => null
                };
            }, TimeSpan.FromSeconds(2)).Result;
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
                        var cType = GetSafeProperty(() => child.ControlType);
                        var cName = GetSafeProperty(() => child.Name);
                        var cId = GetSafeProperty(() => child.AutomationId);

                        if (cType == ControlType.Pane.ToString() &&
                            string.IsNullOrEmpty(cName) &&
                            string.IsNullOrEmpty(cId))
                        {
                            continue;
                        }

                        node.Children.Add(BuildSimplifiedTree(child, currentDepth + 1, maxDepth));
                    }
                }
                catch { }
            }
            return node;
        }

        private string GetSafeProperty(Func<object> propertyGetter)
        {
            try
            {
                var result = propertyGetter();
                return result?.ToString() ?? "";
            }
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
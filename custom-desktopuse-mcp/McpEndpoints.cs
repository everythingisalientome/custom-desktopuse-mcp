using System.ComponentModel;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace DesktopMcpServer
{
    [McpServerToolType]
    public static class McpEndpoints
    {
        // Tool: LaunchApp
        [McpServerTool(Name = "LaunchApp")]
        [Description("Launches a local application executable.")]
        public static string LaunchApp(string path, DesktopAutomationService service)
        {
            return service.LaunchApp(path);
        }

        // Tool: CloseApp
        [McpServerTool(Name = "CloseApp")]
        [Description("Closes an application. Provide 'windowName' (Title or Process Name) to close specific apps. Leave empty to close the current session app.")]
        public static string CloseApp(string? windowName, DesktopAutomationService service)
        {
            return service.CloseApp(windowName);
        }

        // Tool: GetWindowTree
        [McpServerTool(Name = "GetWindowTree")]
        [Description("Finds a window by Title, ID, or Name and returns a JSON tree. Use 'current' for active window.")]
        public static string GetWindowTree(string fieldName, DesktopAutomationService service)
        {
            return service.GetWindowTree(fieldName);
        }

        // Tool: WaitForElement
        [McpServerTool(Name = "WaitForElement")]
        [Description("Waits for a window or element to appear. Useful for long-loading screens. Default timeout is 20s.")]
        public static string WaitForElement(string fieldName, string? windowName, int timeoutSeconds, DesktopAutomationService service)
        {
            // If agent sends 0 or negative, default to 20s
            if (timeoutSeconds <= 0) timeoutSeconds = 20; 
            return service.WaitForElement(fieldName, windowName, timeoutSeconds);
        }

        // Tool: ClickElement
        [McpServerTool(Name = "ClickElement")]
        [Description("Finds an element by ID, Name, Class, or Role and clicks it.")]
        public static string ClickElement(string fieldName, string? windowName, DesktopAutomationService service)
        {
            return service.ClickElement(fieldName, windowName);
        }

        // Tool: RightClickElement
        [McpServerTool(Name = "RightClickElement")]
        [Description("Finds an element by ID, Name, Class, or Role and right-clicks it.")]
        public static string RightClickElement(string fieldName, string? windowName, DesktopAutomationService service)
        {
            return service.RightClickElement(fieldName, windowName);
        }

        // Tool: WriteText
        [McpServerTool(Name = "WriteText")]
        [Description("Writes text to a field. Optionally sends special keys (e.g. '^a{DELETE}') first.")]
        public static string WriteText(string fieldName, string value, string? specialKeys, string? windowName, DesktopAutomationService service)
        {
            return service.WriteText(fieldName, value, specialKeys, windowName);
        }

        // Tool: GetText
        [McpServerTool(Name = "GetText")]
        [Description("Reads text from an element found by ID, Name, Class, or Role.")]
        public static string GetText(string fieldName, string? windowName, DesktopAutomationService service)
        {
            return service.GetText(fieldName, windowName);
        }

        // Tool: SendKeys (Human-Like)
        [McpServerTool(Name = "SendKeys")]
        [Description("Types text slowly, character-by-character with random delays, to simulate a human user.")]
        public static string SendKeys(string fieldName, string value, string? windowName, DesktopAutomationService service)
        {
            return service.SendKeys(fieldName, value, windowName);
        }

        // Tool: SelectItems
        [McpServerTool(Name = "SelectItems")]
        [Description("Selects items in a dropdown or list. 'fieldName' is the list/combo, 'value' is comma-separated items.")]
        public static string SelectItems(string fieldName, string value, string? windowName, DesktopAutomationService service)
        {
            return service.SelectItems(fieldName, value, windowName);
        }

        // Tool: SetCheckbox
        [McpServerTool(Name = "SetCheckbox")]
        [Description("Sets a checkbox. 'fieldName' is the checkbox, 'value' is 'on', 'true', or 'off'.")]
        public static string SetCheckbox(string fieldName, string value, string? windowName, DesktopAutomationService service)
        {
            return service.SetCheckbox(fieldName, value, windowName);
        }

        // Tool: SelectRadioButton
        [McpServerTool(Name = "SelectRadioButton")]
        [Description("Selects a radio button. 'fieldName' is the button Name/ID.")]
        public static string SelectRadioButton(string fieldName, string value, string? windowName, DesktopAutomationService service)
        {
            return service.SelectRadioButton(fieldName, value, windowName);
        }

        // Tool: SendSpecialKeys
        [McpServerTool(Name = "SendSpecialKeys")]
        [Description("Sends global keystrokes (e.g. '{ENTER}', '^c') to the focused window.")]
        public static string SendSpecialKeys(string specialKeys, DesktopAutomationService service)
        {
            return service.SendSpecialKeys(specialKeys);
        }
    }
}
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

        // Tool: GetWindowTree 
        [McpServerTool(Name = "GetWindowTree")]
        [Description("Finds a window by title and returns a simplified JSON tree.")]
        public static string GetWindowTree(string windowTitle, DesktopAutomationService service)
        {
            return service.GetWindowTree(windowTitle);
        }

        // Tool: ClickElement
        [McpServerTool(Name = "ClickElement")]
        [Description("Finds a UI element and clicks it.")]
        public static string ClickElement(string criteria, string value, DesktopAutomationService service)
        {
            return service.ClickElement(criteria, value);
        }

        // Tool: RightClickElement
        [McpServerTool(Name = "RightClickElement")]
        [Description("Finds a UI element and performs a right-click (context menu).")]
        public static string RightClickElement(string criteria, string value, DesktopAutomationService service)
        {
            return service.RightClickElement(criteria, value);
        }

        // Tool: TypeText (Renamed/Refined)
        [McpServerTool(Name = "TypeText")]
        [Description("Finds an edit field and types text into it.")]
        public static string TypeText(string criteria, string value, string text, DesktopAutomationService service)
        {
            return service.TypeText(criteria, value, text);
        }

        // Tool: GetText
        [McpServerTool(Name = "GetText")]
        [Description("Reads the text content from a UI element.")]
        public static string GetText(string criteria, string value, DesktopAutomationService service)
        {
            return service.GetText(criteria, value);
        }

        // Tool: SendKeys (Global keys)
        [McpServerTool(Name = "SendKeys")]
        [Description("Sends global keystrokes to the currently focused window (e.g., '{ENTER}', '^c').")]
        public static string SendKeys(string keys, DesktopAutomationService service)
        {
            return service.SendKeys(keys);
        }

        // Tool: SelectItems
        [McpServerTool(Name = "SelectItems")]
        [Description("Selects one or more items from a dropdown, combo box, or list. Provide multiple items separated by commas.")]
        public static string SelectItems(string criteria, string value, string items, DesktopAutomationService service)
        {
            return service.SelectItems(criteria, value, items);
        }

        // Tool: SetCheckbox
        [McpServerTool(Name = "SetCheckbox")]
        [Description("Sets a checkbox to checked (true/on) or unchecked (false/off).")]
        public static string SetCheckbox(string criteria, string value, string state, DesktopAutomationService service)
        {
            return service.SetCheckbox(criteria, value, state);
        }

        // Tool: SelectRadioButton
        [McpServerTool(Name = "SelectRadioButton")]
        [Description("Selects a specific radio button option.")]
        public static string SelectRadioButton(string criteria, string value, DesktopAutomationService service)
        {
            return service.SelectRadioButton(criteria, value);
        }

        // Tool: CloseApp
        [McpServerTool(Name = "CloseApp")]
        [Description("Closes the currently running application.")]
        public static string CloseApp(DesktopAutomationService service)
        {
            return service.CloseApp();
        }
    }
}
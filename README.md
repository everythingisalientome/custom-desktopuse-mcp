# Desktop Automation MCP Server

A Model Context Protocol (MCP) server built with C# and .NET 8. This server enables AI agents (like GitHub Copilot or Claude Desktop) to view, interact with, and automate Windows desktop applications using standard UI Automation.

## Features & Tools

This server exposes the following tools to the AI:

* **App Management**
    * `LaunchApp`: Starts an application (e.g., `calc.exe`).
    * `CloseApp`: Closes the currently running application.
* **Discovery**
    * `GetWindowTree`: Scans the current window and returns a JSON tree of UI elements (Buttons, Inputs, IDs).
* **Interaction**
    * `ClickElement`: Clicks a button or element.
    * `RightClickElement`: Opens context menus.
    * `WriteText`: Enters text into fields.
    * `GetText`: Reads values from displays or text boxes.
    * `SendKeys`: Sends global keyboard shortcuts (e.g., `{ENTER}`, `^c`).
    * `SelectItems`: Selects items from Dropdowns (ComboBox) or Lists (ListBox).
    * `SetCheckbox`: Checks or unchecks boxes.
    * `SelectRadioButton`: Selects a specific radio option.

## Prerequisites

* **OS:** Windows 10/11
* **Framework:** .NET 8.0 SDK

## Installation & Setup

1.  **Build the Project:**
    Open a terminal in the project folder and run:
    ```powershell
    dotnet build -c Debug
    ```

2.  **Get the Executable Path:**
    Note the full path to the `.exe` generated (usually in `bin\Debug\net8.0-windows\`).

## Configuration

### For VS Code (GitHub Copilot)
Create or update `.vscode/mcp.json` in your workspace:

```json
{
    "mcpServers": {
        "desktop-automation": {
            "command": "C:\\Path\\To\\Your\\Project\\bin\\Debug\\net8.0-windows\\DesktopMcpServer.exe",
            "args": []
        }
    }
}
```
Ideas Help: 
1. https://www.youtube.com/watch?v=MKD-sCZWpZQ
2. https://learn.microsoft.com/en-us/nuget/concepts/nuget-mcp-server
3. https://github.com/FlaUI/FlaUI
4. https://github.com/microsoft/WPF-Samples/tree/main
using System;
using DesktopMcpServer;

namespace TesterApp
{
    class TesterProgram
    {
        static void Main(string[] args)
        {
            string appPath = @"D:\VisualStudioWrkSpce\2022\DemoApps\ExpenseIt\ExpenseItDemo\bin\Debug\net8.0-windows\ExpenseIt9.exe";

            DesktopAutomationService service = new DesktopAutomationService();

            service.LaunchApp(appPath);
            var jsonTree = service.GetWindowTree("ExpenseIt Standalone");
            Console.WriteLine(jsonTree);

            //for WPF applications the field Name is usually the x:Name property is treated as AutomationId in FlaUI so we need to pass as ID
            service.ClickElement("id", "emailTextBox");
            service.TypeText("id", "emailTextBox", "preet.panda@outlook.com");

            service.ClickElement("id", "employeeNumberTextBox");
            service.TypeText("id", "employeeNumberTextBox", "12345");

            service.SelectItems("id", "costCenterTextBox", "Marketing");

            service.SelectRadioButton("name", "CSG");

            service.ClickElement("id", "createExpenseReportButton");
            service.Dispose();
        }
    }
}
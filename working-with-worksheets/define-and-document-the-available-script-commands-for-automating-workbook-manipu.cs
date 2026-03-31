using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ScriptCommandsDemo
{
    // Simple DTO to hold a script command and its description
    public class ScriptCommandInfo
    {
        public string Command { get; }
        public string Description { get; }

        public ScriptCommandInfo(string command, string description)
        {
            Command = command;
            Description = description;
        }
    }

    // Central place that defines and documents the script commands
    public static class WorkbookScriptCommands
    {
        // Collection of known header/footer script commands
        public static readonly List<ScriptCommandInfo> Commands = new List<ScriptCommandInfo>
        {
            new ScriptCommandInfo("&D", "Insert current date"),
            new ScriptCommandInfo("&T", "Insert current time"),
            new ScriptCommandInfo("&\"font\",size", "Set font name and size for following text"),
            new ScriptCommandInfo("&[Page] of &N", "Insert current page number and total page count"),
            new ScriptCommandInfo("&L", "Insert left-aligned text"),
            new ScriptCommandInfo("&C", "Insert centered text"),
            new ScriptCommandInfo("&R", "Insert right-aligned text"),
            // Add additional commands as needed
        };

        // Prints the list of commands to the console
        public static void PrintCommands()
        {
            Console.WriteLine("Available Workbook Script Commands:");
            foreach (var cmd in Commands)
            {
                Console.WriteLine($"{cmd.Command} - {cmd.Description}");
            }
            Console.WriteLine();
        }

        // Demonstrates how to parse a header/footer script using PageSetup.GetCommands
        public static void ShowHeaderFooterCommands(string script)
        {
            // Lifecycle: create a workbook instance
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Use the built‑in method to retrieve all commands from the script
            HeaderFooterCommand[] parsedCommands = worksheet.PageSetup.GetCommands(script);

            Console.WriteLine($"Parsing script: \"{script}\"");
            foreach (HeaderFooterCommand cmd in parsedCommands)
            {
                // cmd.Type is an enum indicating the command type; cmd.Text contains any literal text
                Console.WriteLine($"Type: {cmd.Type}, Text: \"{cmd.Text}\"");
            }

            // Optional: save the temporary workbook (demonstrates the save lifecycle rule)
            workbook.Save("TempWorkbook.xlsx", SaveFormat.Xlsx);
            workbook.Dispose();
        }
    }

    class Program
    {
        static void Main()
        {
            // Document the available script commands
            WorkbookScriptCommands.PrintCommands();

            // Example header/footer script that uses several commands
            string headerScript = "&D&\"Arial\"&12 Sample Header";

            // Parse and display the individual commands from the script
            WorkbookScriptCommands.ShowHeaderFooterCommands(headerScript);
        }
    }
}
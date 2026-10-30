// Title: C# macro to enumerate all worksheets in an Aspose.Cells workbook and log each sheet name with a custom Logger class
// Common Searches: how to enumerate worksheets in an Aspose.Cells workbook using C# | C# example for logging each Excel sheet name with Aspose.Cells | using a static logger to print worksheet names while iterating over Aspose.Cells worksheets | Aspose.Cells macro to list all sheet names in the console | C# code to loop through all worksheets in a workbook and output their names with Aspose.Cells
// Tags: worksheet enumeration Aspose.Cells C# | console logging sheet names Aspose.Cells | static logger class C# Aspose.Cells | Aspose.Cells workbook sheet iteration example | macro to list Excel sheets .NET

using System;
using Aspose.Cells;

namespace AsposeCellsMacroExample
{
    // Simple logging module
    // The example defines a static Logger class that writes messages to the console, creates a Workbook, adds three worksheets, then iterates over the workbook's Worksheets collection and logs each worksheet's Name using the Logger.
    public static class Logger
    {
        // Logs a message to the console (could be extended to log to a file)
        public static void Log(string message)
        {
            Console.WriteLine(message);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook(); // using the create rule

            // Example: add a few worksheets for demonstration
            workbook.Worksheets[0].Name = "Sheet1";
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Iterate through all worksheets and log each sheet name using the Logger module
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Logger.Log($"Worksheet name: {sheet.Name}");
            }

            // Save the workbook if needed (using the save rule)
            // workbook.Save("Output.xlsx");
        }
    }
}

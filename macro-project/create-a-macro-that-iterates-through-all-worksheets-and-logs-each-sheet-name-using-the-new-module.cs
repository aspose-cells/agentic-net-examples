// Title: Iterate all worksheets and log each sheet name using a custom Logger module with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# iterate through all worksheets and print sheet names to console | How to use a static logger to output Excel worksheet names in a .NET application | Example of enumerating worksheets in Aspose.Cells and logging each name | C# code to log workbook sheet names using a custom logger module with Aspose.Cells
// Tags: log worksheet names Aspose.Cells C# | enumerate worksheets Aspose.Cells .NET | custom static logger module C# console | save workbook as Output.xlsx Aspose.Cells | create and add worksheets Aspose.Cells example

using Aspose.Cells;
using System;

namespace AsposeCellsMacroExample
{
    // Simple logger module
    // The sample creates a Workbook, adds three worksheets, iterates over workbook.Worksheets, logs each sheet's Name to the console via a static Logger.Log method, and saves the file as Output.xlsx.
    public static class Logger
    {
        // Logs a message (here to the console)
        public static void Log(string message)
        {
            Console.WriteLine(message);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample worksheets (optional, for demonstration)
            workbook.Worksheets[0].Name = "Sheet1";
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Iterate through all worksheets and log each sheet name using the Logger module
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Logger.Log($"Worksheet name: {sheet.Name}");
            }

            // Save the workbook (optional)
            workbook.Save("Output.xlsx");
        }
    }
}

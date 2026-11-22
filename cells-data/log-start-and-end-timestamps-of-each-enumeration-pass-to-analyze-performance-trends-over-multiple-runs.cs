// Title: How to log start and end timestamps for each worksheet and cell enumeration pass using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells log timestamps for worksheet enumeration passes | measure performance of worksheet and cell loops with Aspose.Cells .NET | record start and end time of each enumeration pass in an Excel workbook using Aspose.Cells | Aspose.Cells performance logging example for multiple enumeration passes | write UTC timestamps to a log file during Aspose.Cells workbook processing
// Tags: Aspose.Cells enumeration timing | worksheet iteration performance logging C# | cell enumeration duration Aspose.Cells | UTC timestamp logging Aspose.Cells workbook | performance measurement Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPerformanceLog
{
    // The sample loads an Excel workbook with Aspose.Cells, clears a log file, iterates through all worksheets and then through all cells of the first worksheet, and records UTC timestamps for the start and end of each enumeration pass to a text file while also printing them to the console.
    class Program
    {
        // Path to the Excel file to be processed
        private const string InputFilePath = @"C:\Data\Sample.xlsx";

        // Path to the log file where timestamps will be recorded
        private const string LogFilePath = @"C:\Data\PerformanceLog.txt";

        static void Main()
        {
            // Ensure the log file exists and clear previous content
            File.WriteAllText(LogFilePath, string.Empty);

            // Load the workbook (creation/loading must follow provided lifecycle rules)
            Workbook workbook = new Workbook(InputFilePath);

            // Example enumeration pass: iterate through all worksheets
            LogTimestamp("Enumeration Pass 1 - Start");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Simulate some processing per worksheet (e.g., reading cell values)
                var cellValue = sheet.Cells["A1"].Value;
                // (Processing logic can be expanded as needed)
            }
            LogTimestamp("Enumeration Pass 1 - End");

            // Second enumeration pass: iterate through all cells in the first worksheet
            LogTimestamp("Enumeration Pass 2 - Start");
            Worksheet firstSheet = workbook.Worksheets[0];
            foreach (Cell cell in firstSheet.Cells)
            {
                // Simulate processing each cell
                var value = cell.Value;
            }
            LogTimestamp("Enumeration Pass 2 - End");

            // Additional enumeration passes can be added here following the same pattern

            // Save the workbook if any modifications were made (must use provided save rule)
            // workbook.Save(@"C:\Data\Sample_Processed.xlsx");
        }

        /// <param name="message">Description of the event being logged.</param>
        private static void LogTimestamp(string message)
        {
            string entry = $"{DateTime.UtcNow:O} - {message}";
            Console.WriteLine(entry);
            File.AppendAllLines(LogFilePath, new[] { entry });
        }
    }
}

// Title: Copy a chart sheet to a new worksheet while preserving controls and formatting with Aspose.Cells for .NET (C#)
// AI Prompts: Provide C# code that loads a workbook, uses Worksheets.AddCopy to clone the chart sheet 'ChartSheet1', renames the new sheet to 'ChartSheetCopy', and saves the file, ensuring all chart objects remain intact. | Write a C# console program that validates the existence of a source chart sheet, copies it to a new worksheet, keeps all attached controls and formatting, and includes try‑catch error handling with Aspose.Cells. | Generate a minimal Aspose.Cells example in C# that demonstrates cloning a chart sheet, preserving its formatting, and exporting the result to 'output.xlsx'.
// Common Searches: Aspose.Cells C# how to duplicate a chart sheet with formatting | copy chart sheet to another worksheet using Aspose.Cells .NET | preserve chart controls when copying chart sheet in Aspose.Cells | example of Worksheets.AddCopy for chart sheets in C#
// Tags: Worksheets.AddCopy chart sheet operation | chart sheet cloning Aspose.Cells | retain chart formatting during sheet copy | Aspose.Cells preserve chart controls | C# chart sheet copy example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads 'source.xlsx', locates the chart sheet named 'ChartSheet1', clones it to a new sheet called 'ChartSheetCopy' using Worksheets.AddCopy, preserves all chart objects and formatting, and saves the workbook as 'output.xlsx' with basic error handling.
class ChartCopyExample
{
    static void Main()
    {
        try
        {
            const string sourceFile = "source.xlsx";
            const string outputFile = "output.xlsx";

            // Verify source workbook exists
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file \"{sourceFile}\" not found.");
                return;
            }

            // Load the existing workbook that contains the source chart sheet
            Workbook workbook = new Workbook(sourceFile);

            // Find the source chart sheet by name
            Worksheet sourceChartSheet = workbook.Worksheets["ChartSheet1"];
            if (sourceChartSheet == null)
            {
                Console.WriteLine("Source chart sheet \"ChartSheet1\" not found.");
                return;
            }

            // Get the index of the source chart sheet
            int sourceChartSheetIndex = workbook.Worksheets.IndexOf(sourceChartSheet);

            // Determine a name for the destination chart sheet
            string destChartSheetName = "ChartSheetCopy";

            // Copy the source chart sheet
            int destIndex = workbook.Worksheets.AddCopy(sourceChartSheetIndex);
            Worksheet destChartSheet = workbook.Worksheets[destIndex];
            destChartSheet.Name = destChartSheetName;

            // Save the workbook with the copied chart sheet
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

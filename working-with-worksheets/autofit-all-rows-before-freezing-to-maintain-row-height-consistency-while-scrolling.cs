// Title: How to auto‑fit all rows and then freeze the top row in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing .xlsx workbook with Aspose.Cells, determines the used range, calls AutoFitRows for every row, freezes the first row while keeping all columns scrollable, and saves the modified file. | Show a C# example that checks for the input file, uses Worksheet.AutoFitRows(0, lastRow) and Worksheet.FreezePanes(1, 0, lastRow, lastColumn) to maintain row heights when freezing panes, and includes basic exception handling. | Provide a robust Aspose.Cells snippet that calculates MaxDataRow/MaxDataColumn, applies AutoFitRows before FreezePanes, and writes the result to a new workbook with proper error messages.
// Common Searches: Aspose.Cells C# auto fit rows then freeze first row in Excel | C# code to auto adjust row heights before applying FreezePanes with Aspose.Cells | How to keep row height consistent when freezing panes using Aspose.Cells for .NET | Determine used range and auto fit rows in Aspose.Cells before freezing header row | Freeze top row while preserving column visibility Aspose.Cells .NET example
// Tags: auto-fit rows Aspose.Cells .NET | freeze top row worksheet Aspose.Cells | calculate used range Aspose.Cells | preserve row height freeze panes .NET | load and save workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads input.xlsx, computes the used rows and columns, auto‑fits all rows to their content, freezes the first row while keeping all columns visible, and saves the result to output.xlsx, with file‑existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the number of rows and columns that contain data
                int totalRows = sheet.Cells.MaxDataRow + 1;       // +1 because rows are zero‑based
                int totalColumns = sheet.Cells.MaxDataColumn + 1; // +1 because columns are zero‑based

                // Auto‑fit all rows to match their content height
                sheet.AutoFitRows(0, totalRows);

                // Freeze the top row (row index 1) while keeping all columns visible
                // FreezePanes(row, column, totalRows, totalColumns)
                sheet.FreezePanes(1, 0, totalRows, totalColumns);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

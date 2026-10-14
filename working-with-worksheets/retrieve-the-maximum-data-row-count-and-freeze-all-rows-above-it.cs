// Title: Freeze all rows above the last populated row in each worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, finds the last row containing data on each worksheet via Cells.MaxDataRow, and uses FreezePanes to lock all rows above that row. | Show how to iterate through every worksheet in a workbook and apply the FreezePanes overload (rows, columns, totalRows, totalColumns) to freeze rows up to the maximum data row.
// Common Searches: Aspose.Cells C# freeze rows up to the last non‑empty row in each sheet | How to use Cells.MaxDataRow with FreezePanes in a .NET Excel workbook | Programmatically freeze top rows based on data length using Aspose.Cells for .NET
// Tags: freeze panes based on Cells.MaxDataRow Aspose.Cells | determine last data row worksheet Aspose.Cells .NET | apply FreezePanes to all worksheets C# | Aspose.Cells workbook row freezing example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads a workbook, loops through each worksheet, retrieves the last row with data via Cells.MaxDataRow, freezes rows 0‑maxDataRow using FreezePanes, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; if not, create an empty workbook.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file not found: {inputPath}. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Iterate through each worksheet and freeze rows up to the last data row.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int maxDataRow = sheet.Cells.MaxDataRow; // 0‑based index of last row with data

                    if (maxDataRow >= 0)
                    {
                        // Freeze rows 0..maxDataRow (no columns). 
                        // Use overload with totalRows and totalColumns.
                        sheet.FreezePanes(maxDataRow + 1, 0, maxDataRow + 1, 0);
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

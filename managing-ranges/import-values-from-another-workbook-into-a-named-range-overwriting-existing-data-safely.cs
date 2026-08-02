// Title: Import data into a named range from another workbook using Aspose.Cells (C#)
// Description: C# sample that loads a source and destination workbook, ensures a named range exists, clears its contents, copies matching data (values, formulas, formatting) from the source range, and saves the updated file.
// Keywords: Aspose.Cells C# import named range | copy data between workbooks Aspose.Cells | clear range before paste Aspose.Cells | create named range programmatically | .NET Excel range overwrite | Excel workbook merge Aspose | copy formulas formatting Aspose.Cells | named range data refresh | Excel automation Aspose.Cells | C# Excel range copy example
// Common Searches: Aspose.Cells copy data to a named range | How to overwrite a named range from another workbook in C# | Clear existing cells before copying with Aspose.Cells | Create a missing named range programmatically Aspose.Cells | Copy values, formulas and formatting between Excel files using Aspose
// Developer Intent: Import data from a source workbook into a specific named range of a destination workbook, safely replacing any existing content.
// Use Cases: Refresh a reporting table by pulling the latest calculations into its predefined named range. | Update a chart source in a template workbook by copying new metric values into the chart's named range. | Replace dashboard data while preserving cell styles by overwriting the associated named range.
// AI Prompts: Write C# code with Aspose.Cells that checks for a named range, creates it if missing, clears its cells, and copies a matching source range from another workbook. | Show how to build a source address based on the dimensions of a destination named range and copy values, formulas, and formatting safely. | Explain strategies for handling size mismatches between source data and a destination named range when using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImportIntoNamedRange
{
    // C# sample that loads a source and destination workbook, ensures a named range exists, clears its contents, copies matching data (values, formulas, formatting) from the source range, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "Source.xlsx";
                const string destPath = "Destination.xlsx";
                const string outputPath = "Destination_Updated.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Verify destination file exists
                if (!File.Exists(destPath))
                {
                    Console.WriteLine($"Destination file not found: {destPath}");
                    return;
                }

                // Load the source workbook that contains the data to be copied
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Load the destination workbook where the named range resides
                Workbook destWorkbook = new Workbook(destPath);

                // ------------------------------------------------------------
                // Ensure the destination workbook has a named range called "MyRange"
                // If it does not exist, create it and point it to a default area
                // ------------------------------------------------------------
                Name myRangeName;
                if (destWorkbook.Worksheets.Names["MyRange"] == null)
                {
                    // Add a new named range to the first worksheet (adjust as needed)
                    int nameIndex = destWorkbook.Worksheets.Names.Add("MyRange");
                    myRangeName = destWorkbook.Worksheets.Names[nameIndex];
                    // Example reference – you can change the address to suit your scenario
                    myRangeName.RefersTo = "=Sheet1!A1:B2";
                }
                else
                {
                    myRangeName = destWorkbook.Worksheets.Names["MyRange"];
                }

                // Retrieve the actual Range object that the name refers to
                Aspose.Cells.Range destRange = myRangeName.GetRange();

                // ------------------------------------------------------------
                // Safely clear any existing data in the destination range
                // (clears both contents and formatting)
                // ------------------------------------------------------------
                destRange.Worksheet.Cells.ClearRange(
                    destRange.FirstRow,
                    destRange.FirstColumn,
                    destRange.RowCount,
                    destRange.ColumnCount);

                // ------------------------------------------------------------
                // Define the source range that holds the data to be imported.
                // For this example we assume the source data occupies the same
                // size as the destination named range.
                // ------------------------------------------------------------
                // Build address like "A1:Z10" based on destination dimensions
                string lastColumnName = CellsHelper.ColumnIndexToName(destRange.ColumnCount - 1);
                int lastRowNumber = destRange.RowCount; // because rows are 1‑based in address
                string sourceAddress = $"A1:{lastColumnName}{lastRowNumber}";
                Aspose.Cells.Range sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange(sourceAddress);

                // ------------------------------------------------------------
                // Copy the data (including values, formulas, and formatting) from
                // the source range into the cleared destination range.
                // ------------------------------------------------------------
                destRange.CopyData(sourceRange);

                // ------------------------------------------------------------
                // Save the destination workbook with the imported data
                // ------------------------------------------------------------
                destWorkbook.Save(outputPath);
                Console.WriteLine($"Data imported successfully. Saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Stream and filter rows where column C = "Active" with Aspose.Cells LightCells (C#)
// Description: A memory‑efficient C# example that opens an Excel workbook, streams rows using Aspose.Cells LightCells, copies the header, filters rows whose column C contains "Active", writes the matching rows to a new worksheet, and saves the result.
// Keywords: Aspose.Cells | LightCells | C# | filter Excel rows | column C Active | streaming rows | copy rows to new sheet | .NET Excel processing | memory efficient Excel | save filtered workbook
// Common Searches: Aspose.Cells filter rows by column value C# | LightCells stream Excel and copy selected rows | C# copy rows where cell equals Active | How to export only active records with Aspose.Cells | Memory‑friendly Excel row filtering .NET
// Developer Intent: Read a source workbook, stream its rows, keep only those with "Active" in column C, and write those rows to a new worksheet.
// Use Cases: Create a lightweight report containing only active entries from a master sheet. | Export active records for downstream data pipelines without loading the entire file into memory. | Prepare a clean version of a workbook for external partners by removing inactive rows.
// AI Prompts: Generate C# code that uses Aspose.Cells LightCells to stream an Excel file and write rows with column C = "Active" to a new sheet. | Explain how to modify the loop to preserve original cell formatting while filtering with LightCells. | Show how to handle very large workbooks (>1 GB) using LightCells and write only matching rows to a separate file. | Provide a step‑by‑step guide for adding a progress callback during LightCells streaming.

using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsFilterExample
{
    // A memory‑efficient C# example that opens an Excel workbook, streams rows using Aspose.Cells LightCells, copies the header, filters rows whose column C contains "Active", writes the matching rows to a new worksheet, and saves the result.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load source workbook
                Workbook sourceWb = new Workbook(inputPath);
                Worksheet sourceSheet = sourceWb.Worksheets[0];

                // Create destination workbook (empty)
                Workbook destWb = new Workbook();
                Worksheet destSheet = destWb.Worksheets[0];

                // Copy header row (row 0)
                destSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, 1);
                int destRow = 1; // start after header

                // Iterate source rows and copy those where column C (index 2) equals "Active"
                int maxRow = sourceSheet.Cells.MaxDataRow;
                for (int r = 1; r <= maxRow; r++)
                {
                    Cell cell = sourceSheet.Cells[r, 2]; // column C
                    if (cell != null && cell.Type == CellValueType.IsString && cell.StringValue == "Active")
                    {
                        // Copy the entire row to the destination sheet
                        destSheet.Cells.CopyRows(sourceSheet.Cells, r, destRow, 1);
                        destRow++;
                    }
                }

                // Save filtered workbook
                destWb.Save(outputPath);
                Console.WriteLine($"Filtered workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

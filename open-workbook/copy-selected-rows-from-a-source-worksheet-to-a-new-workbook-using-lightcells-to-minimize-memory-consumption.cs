// Title: Copy specific rows from a worksheet to a new workbook with Aspose.Cells LightCells in C#
// AI Prompts: Write C# code that uses Aspose.Cells LightCells to extract rows 1, 3, and 4 from a source worksheet and write them to a new workbook while preserving cell values and styles. | Show how to implement memory‑efficient row copying in Aspose.Cells by iterating with LightCells and saving the selected rows to a separate Excel file. | Refactor the provided example to leverage LightCells for selective row export and reduce overall memory consumption.
// Common Searches: Aspose.Cells C# copy selected rows to new Excel file using LightCells | How to export only certain rows from a worksheet with Aspose.Cells without loading the whole sheet | Memory efficient row extraction Aspose.Cells LightCells example C# | Preserve formatting when copying specific rows with Aspose.Cells | Select rows by index and save to another workbook Aspose.Cells C#
// Tags: lightcells selective row copy C# | aspocells copy rows to new workbook | excel row extraction memory efficient | preserve cell style Aspose.Cells | copy rows without loading full sheet

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program loads source.xlsx, uses LightCells to iterate only over rows 1, 3, and 4, copies each cell’s value and style to a new workbook, and saves the result as selected_rows.xlsx, minimizing memory usage.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "selected_rows.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(sourcePath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Create a new workbook for the selected rows
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Zero‑based indices of rows to copy (e.g., rows 2, 4, 5 in Excel are 1,3,4)
            int[] rowsToCopy = new int[] { 1, 3, 4 };

            // Determine the maximum column that contains data in the source sheet
            int maxColumn = srcSheet.Cells.MaxDataColumn;

            int destRowIndex = 0; // Destination row pointer

            foreach (int srcRowIndex in rowsToCopy)
            {
                // Copy each cell in the selected row
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell srcCell = srcSheet.Cells[srcRowIndex, col];
                    if (srcCell == null) continue;

                    Cell destCell = destSheet.Cells[destRowIndex, col];
                    destCell.Value = srcCell.Value;

                    // Preserve style
                    Style style = srcCell.GetStyle();
                    destCell.SetStyle(style);
                }

                destRowIndex++; // Move to next row in destination sheet
            }

            // Save the result workbook
            destWorkbook.Save(destPath);
            Console.WriteLine($"Selected rows saved to: {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

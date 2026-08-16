// Title: C# Aspose.Cells: Measure Excel file size before and after RemoveUnusedStyles
// Description: A concise example that creates a workbook, assigns a distinct style to each cell, deletes rows to leave orphaned styles, saves the file, records its byte size, invokes Workbook.RemoveUnusedStyles() to purge unused styles, saves the cleaned file, and then reports the absolute and percentage size reduction. Ideal for developers who need to quantify the impact of style cleanup on Excel file size.
// Keywords: Aspose.Cells RemoveUnusedStyles | C# Excel file size measurement | Workbook size optimization .NET | Excel style cleanup impact | Aspose.Cells performance tuning | file size reduction Aspose.Cells | C# measure Excel bytes | remove orphaned styles Aspose
// Common Searches: how to check Excel file size with Aspose.Cells C# | remove unused styles and see size difference Aspose.Cells | Workbook.RemoveUnusedStyles effect on file size | C# code to compare Excel sizes before and after cleanup | Aspose.Cells reduce workbook size by deleting styles
// Developer Intent: Find out how much an Excel workbook shrinks when unused cell styles are removed using Aspose.Cells for .NET.
// Use Cases: Validate that style cleanup lowers the size of large, program‑generated reports before distribution. | Integrate size‑reduction logging into a batch process that processes thousands of Excel files. | Generate audit reports that record the byte and percentage savings achieved by removing orphaned styles.
// AI Prompts: Generate C# code with Aspose.Cells that logs workbook size before and after calling RemoveUnusedStyles and outputs the percentage reduction. | Explain how Workbook.RemoveUnusedStyles works internally and which parts of the XLSX package are affected. | Suggest additional Aspose.Cells techniques (e.g., compressing images, removing empty rows/columns) to further reduce Excel file size.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStyleSizeDemo
{
    // A concise example that creates a workbook, assigns a distinct style to each cell, deletes rows to leave orphaned styles, saves the file, records its byte size, invokes Workbook.RemoveUnusedStyles() to purge unused styles, saves the cleaned file, and then reports the absolute and percentage size reduction. Ideal for developers who need to quantify the impact of style cleanup on Excel file size.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data with distinct styles
            for (int i = 0; i < 10; i++)
            {
                Cell cell = cells[i, 0];
                cell.PutValue($"Item {i + 1}");

                // Create a new style for each cell
                Style style = wb.CreateStyle();
                style.Font.Name = "Arial";
                style.Font.Size = 10 + i;
                style.Font.IsBold = i % 2 == 0;
                cell.SetStyle(style);
            }

            // Delete some rows to leave unused styles in the workbook
            sheet.Cells.DeleteRows(5, 5);

            // Save the workbook before removing unused styles
            string beforePath = "BeforeRemoveUnusedStyles.xlsx";
            wb.Save(beforePath);

            // Measure file size before cleanup
            long sizeBefore = new FileInfo(beforePath).Length;
            Console.WriteLine($"File size before removing unused styles: {sizeBefore} bytes");

            // Remove all unused styles
            wb.RemoveUnusedStyles();

            // Save the workbook after cleanup
            string afterPath = "AfterRemoveUnusedStyles.xlsx";
            wb.Save(afterPath);

            // Measure file size after cleanup
            long sizeAfter = new FileInfo(afterPath).Length;
            Console.WriteLine($"File size after removing unused styles: {sizeAfter} bytes");

            // Display reduction information
            long reduction = sizeBefore - sizeAfter;
            double percent = sizeBefore > 0 ? (double)reduction / sizeBefore * 100 : 0;
            Console.WriteLine($"Size reduction: {reduction} bytes ({percent:F2}%)");
        }
    }
}

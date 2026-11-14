// Title: How to set a minimum row height and then auto‑fit a row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set minimum row height before AutoFitRow | C# enforce lower bound on row height then auto‑fit with Aspose.Cells | Combine SetRowHeight and AutoFitRow to keep row height at least 30 points in Aspose.Cells | Prevent AutoFitRow from shrinking row height in Aspose.Cells .NET | Set row height limit and auto‑adjust row size using Aspose.Cells for C#
// Tags: setrowheight with autofitrow Aspose.Cells | minimum row height enforcement .NET | auto‑fit row with height constraint C# | Aspose.Cells row height lower bound | excel row height auto‑adjust Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, writes long text to cell A2, sets a minimum row height of 30 points for row 2, calls AutoFitRow to expand the height if needed, verifies the final height does not fall below the minimum, and saves the file as 'CombinedRowHeightAutoFit.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Example data in row 2 (index 1)
                cells["A2"].PutValue("This is a long text that may require a taller row after auto‑fit.");

                // Define the row index (zero‑based) and the minimum height (in points)
                int rowIndex = 1;               // Row 2 in Excel
                double minHeight = 30.0;        // Minimum row height

                // Set the minimum row height first
                sheet.Cells.SetRowHeight(rowIndex, minHeight);

                // Auto‑fit the row; this will increase the height if the content needs more space
                sheet.AutoFitRow(rowIndex);

                // Ensure the row height is not less than the minimum (in case AutoFit reduced it)
                if (sheet.Cells.GetRowHeight(rowIndex) < minHeight)
                {
                    sheet.Cells.SetRowHeight(rowIndex, minHeight);
                }

                // Save the workbook to a file
                string outputPath = "CombinedRowHeightAutoFit.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Create an Excel file with a count subtotal for column B (rows 5‑200) using outline grouping in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add count subtotal with outline grouping for a specific row range | How to generate a subtotal that counts values in column B rows 5 to 200 using Aspose.Cells | C# code to apply outline grouping and count subtotal on Excel column with Aspose.Cells | Using CellArea.Subtotal to count entries in a range with Aspose.Cells .NET | Create grouped rows and subtotal summary below in Excel via Aspose.Cells C#
// Tags: Aspose.Cells count subtotal outline grouping | CellArea.Subtotal method C# example | Excel subtotal by column using Aspose.Cells | Generate random data column B Aspose.Cells | C# create grouped rows in Excel workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // The example creates a new workbook, fills column B rows 5‑200 with random integers, defines a CellArea covering rows 5‑200 and columns A‑B, and applies the Aspose.Cells Subtotal method with the Count function and outline grouping. The subtotal row is placed below each group, column B width is adjusted, and the workbook is saved as SubtotalOutlineDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the output Excel file
                string outputPath = "SubtotalOutlineDemo.xlsx";

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // Populate sample data in column B (index 1) for rows 5‑200
                // ------------------------------------------------------------
                // Row indices are zero‑based, so row 5 is index 4.
                int startDataRow = 4;          // Row 5
                int endDataRow = 199;          // Row 200
                Random rnd = new Random();

                for (int row = startDataRow; row <= endDataRow; row++)
                {
                    // Put a random integer between 1 and 100 in column B
                    cells[row, 1].PutValue(rnd.Next(1, 101));
                }

                // ------------------------------------------------------------
                // Add a subtotal that counts the entries in column B (index 1)
                // for the range rows 5‑200 and place the summary row below the group.
                // ------------------------------------------------------------
                // Define the area that contains the data (columns A‑B)
                CellArea area = CellArea.CreateCellArea(startDataRow, 0, endDataRow, 1);

                // Group by column A (index 0) – even if empty – and subtotal column B (index 1)
                int groupByColumn = 0;
                int[] subtotalColumns = new int[] { 1 };

                // Apply subtotal: count function, do not replace existing subtotals,
                // do not insert page breaks, place summary row below the group.
                cells.Subtotal(area, groupByColumn, ConsolidationFunction.Count, subtotalColumns, false, false, true);

                // ------------------------------------------------------------
                // (Optional) Adjust column width for better visibility
                // ------------------------------------------------------------
                cells.SetColumnWidth(1, 15); // Column B

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

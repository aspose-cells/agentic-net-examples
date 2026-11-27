// Title: Create an Excel subtotal with StdDev on column H and bottom summary using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add standard deviation subtotal grouped by column | How to place subtotal row at the bottom of each group with Aspose.Cells | C# code to calculate StdDev subtotal for a column in Excel using Aspose.Cells | Using Aspose.Cells Subtotal method with ConsolidationFunction.StdDev | Generate Excel file with grouped StdDev subtotal using Aspose.Cells .NET
// Tags: Aspose.Cells Subtotal StdDev C# | grouped subtotal bottom position Aspose.Cells | Excel StdDev subtotal using Aspose.Cells | C# calculate standard deviation subtotal Aspose.Cells | Aspose.Cells Subtotal function example

using System;
using Aspose.Cells;

// The example creates a new workbook, fills column G with categories and column H with random numeric values, defines a range covering these columns, and applies the Subtotal method to group rows by column G while calculating the standard deviation of column H. The subtotal summary line is placed below each group, and the workbook is saved as SubtotalStdDevBottom.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // -----------------------------------------------------------------
            // 2. Populate sample data.
            //    Column G (index 6) will be used for grouping.
            //    Column H (index 7) will contain numeric values on which we
            //    calculate the StdDev subtotal.
            // -----------------------------------------------------------------
            // Header row
            cells[0, 6].PutValue("Category");   // Column G header
            cells[0, 7].PutValue("Amount");     // Column H header

            // Sample data (10 rows)
            string[] categories = { "A", "B", "A", "B", "A", "B", "A", "B", "A", "B" };
            Random rnd = new Random();

            for (int i = 0; i < 10; i++)
            {
                // Category column (G)
                cells[i + 1, 6].PutValue(categories[i]);

                // Numeric column (H) – random values between 10 and 100
                cells[i + 1, 7].PutValue(rnd.Next(10, 101));
            }

            // -----------------------------------------------------------------
            // 3. Define the range that contains the header and the data.
            //    The range starts at row 0 (header) and ends at row 10
            //    (10 data rows + header). Columns are G (6) to H (7).
            // -----------------------------------------------------------------
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 10,      // header + 10 data rows
                StartColumn = 6,
                EndColumn = 7
            };

            // -----------------------------------------------------------------
            // 4. Add a subtotal.
            //    - Group by column G (index 6)
            //    - Apply the StdDev function on column H (index 7)
            //    - Place the summary (subtotal) line at the bottom of each group
            // -----------------------------------------------------------------
            int groupByColumnIndex = 6;                 // Column G
            int[] subtotalColumnIndexes = new int[] { 7 }; // Column H
            bool summaryBelow = true;                   // Bottom position
            bool replaceExisting = false;               // Do not replace existing subtotals
            bool addPageBreaks = false;                 // No page breaks

            cells.Subtotal(
                area,
                groupByColumnIndex,
                ConsolidationFunction.StdDev,
                subtotalColumnIndexes,
                summaryBelow,
                replaceExisting,
                addPageBreaks);

            // -----------------------------------------------------------------
            // 5. Save the workbook.
            // -----------------------------------------------------------------
            string outputPath = "SubtotalStdDevBottom.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

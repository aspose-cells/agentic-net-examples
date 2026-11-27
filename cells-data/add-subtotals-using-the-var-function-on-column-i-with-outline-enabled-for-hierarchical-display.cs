// Title: Add VAR subtotals to column I with outline grouping in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: how to calculate variance subtotal for a column with Aspose.Cells C# | Aspose.Cells C# create subtotal using VAR function and show outline rows | group rows by column H and add VAR subtotal on column I in Excel via Aspose.Cells | enable outline view and place summary rows below data using Aspose.Cells .NET | use CellArea and Subtotal method to add variance subtotal in Aspose.Cells
// Tags: VAR subtotal Aspose.Cells C# | hierarchical outline Aspose.Cells | CellArea range subtotal Aspose.Cells | group by column H Aspose.Cells | summary rows placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, adds sample data with groups in column H and numeric values in column I, defines a CellArea covering these columns, applies a VAR subtotal on column I grouped by column H using the Subtotal method (replacing existing subtotals and placing summary rows below), enables hierarchical outline view, and saves the result as SubtotalVarResult.xlsx.
class Program
{
    static void Main()
    {
        // Path to an optional input workbook. If the file does not exist we will create a sample workbook.
        string inputPath = "SampleData.xlsx";
        Workbook workbook;

        if (File.Exists(inputPath))
        {
            // Load existing workbook
            workbook = new Workbook(inputPath);
        }
        else
        {
            // Create a new workbook with sample data
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // Header row
            ws.Cells["H1"].PutValue("Group");   // Column H (index 7)
            ws.Cells["I1"].PutValue("Value");   // Column I (index 8)

            // Sample data
            string[] groups = { "A", "A", "B", "B", "C", "C", "C" };
            double[] values = { 10, 20, 15, 25, 30, 40, 50 };

            for (int i = 0; i < groups.Length; i++)
            {
                ws.Cells[i + 1, 7].PutValue(groups[i]);   // Column H
                ws.Cells[i + 1, 8].PutValue(values[i]);   // Column I
            }
        }

        // Work with the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range (including headers)
        int lastRow = worksheet.Cells.MaxDataRow;
        // Create a CellArea that spans from H1 to I<lastRow>
        CellArea area = CellArea.CreateCellArea(0, 7, lastRow, 8); // (startRow, startColumn, endRow, endColumn)

        // Subtotal settings:
        // - Group by the first column of the area (column H, index 0 relative to the area)
        // - Use the VAR function on column I (second column of the area, index 1)
        // - Place summary rows below the detail rows
        // - Replace any existing subtotals in the range
        // - Do not insert page breaks between groups
        int[] totalList = new int[] { 1 }; // column I relative index within the area
        worksheet.Cells.Subtotal(
            area,
            0,                                 // group by column H (relative index 0)
            ConsolidationFunction.Var,         // VAR function
            totalList,                         // columns to subtotal
            true,                              // replace existing subtotals
            false,                             // no page breaks
            true                               // summary rows below data
        );

        // Enable outline view for hierarchical display
        worksheet.IsOutlineShown = true;
        // Ensure summary rows appear below the grouped data
        worksheet.Outline.SummaryRowBelow = true;

        // Save the workbook
        string outputPath = "SubtotalVarResult.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Subtotal with VAR function applied successfully. File saved as '{outputPath}'.");
    }
}

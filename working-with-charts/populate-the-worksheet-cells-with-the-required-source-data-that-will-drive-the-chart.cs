// Title: Populate worksheet cells with source data for an Excel chart using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells write header and data rows for chart source range | How to fill an Excel worksheet with string and double arrays for chart creation using Aspose.Cells | Aspose.Cells example to create chart data range in a new workbook | Saving worksheet with chart source data to .xlsx using Aspose.Cells .NET | Initialize workbook and populate cells for Excel chart with Aspose.Cells
// Tags: header row insertion Aspose.Cells C# | array data population worksheet Aspose.Cells | excel chart data range creation Aspose.Cells | save workbook to xlsx Aspose.Cells | initialize worksheet cells for chart Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook, adds a header row, defines category and sales arrays, iterates to write each value into columns A, B, and C of the first worksheet, and saves the file as ChartData.xlsx, providing the data range required for an Excel chart.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Write header row
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Sales Q1");
        sheet.Cells["C1"].PutValue("Sales Q2");

        // Sample source data
        string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
        double[] salesQ1 = { 1200, 1500, 1100, 1800, 1600 };
        double[] salesQ2 = { 1300, 1400, 1150, 1750, 1650 };

        // Populate the worksheet with the data
        for (int i = 0; i < categories.Length; i++)
        {
            // Column A: Category
            sheet.Cells[i + 1, 0].PutValue(categories[i]);

            // Column B: Sales Q1
            sheet.Cells[i + 1, 1].PutValue(salesQ1[i]);

            // Column C: Sales Q2
            sheet.Cells[i + 1, 2].PutValue(salesQ2[i]);
        }

        // Save the workbook (the chart can later reference this data range)
        workbook.Save("ChartData.xlsx");
    }
}

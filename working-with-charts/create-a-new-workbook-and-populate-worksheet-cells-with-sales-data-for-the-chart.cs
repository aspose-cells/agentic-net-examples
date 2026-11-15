// Title: Generate an Excel workbook with monthly sales data using Aspose.Cells for .NET (C#)
// Common Searches: how to write month and sales values to an Excel sheet with Aspose.Cells in C# | Aspose.Cells create worksheet, add header row and fill data from arrays | save workbook as .xlsx after populating sales data using Aspose.Cells .NET | C# code to populate Excel cells with string and double arrays using Aspose.Cells | rename first worksheet and add sales chart data with Aspose.Cells in C#
// Tags: create workbook Aspose.Cells | populate worksheet from string double arrays | write header row Aspose.Cells | save workbook xlsx Aspose.Cells | setup sales data for chart Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, renames the first worksheet to "SalesData", writes a header row for Month and Sales, fills columns A and B with month names and corresponding sales figures from predefined string and double arrays, and saves the file as "SalesData.xlsx" in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet and rename it
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SalesData";

        // Write header row
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        // Sample sales data
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        double[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };

        // Populate worksheet with data
        for (int i = 0; i < months.Length; i++)
        {
            // Column A: Month
            sheet.Cells[i + 2, 0].PutValue(months[i]);
            // Column B: Sales
            sheet.Cells[i + 2, 1].PutValue(sales[i]);
        }

        // Save the workbook to a file
        workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);
    }
}

using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add a worksheet named "SalesData"
        Worksheet sheet = workbook.Worksheets.Add("SalesData");

        // Write header cells
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        // Sample sales data
        string[] months = { "January", "February", "March", "April", "May" };
        double[] sales = { 12000.5, 15000, 13000, 17000.75, 16000 };

        // Populate the worksheet with the data
        for (int i = 0; i < months.Length; i++)
        {
            // Row index is zero‑based; i+1 corresponds to rows 2,3,...
            sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
            sheet.Cells[i + 1, 1].PutValue(sales[i]);    // Column B
        }

        // Save the workbook to a file
        workbook.Save("SalesReport.xlsx", SaveFormat.Xlsx);
    }
}
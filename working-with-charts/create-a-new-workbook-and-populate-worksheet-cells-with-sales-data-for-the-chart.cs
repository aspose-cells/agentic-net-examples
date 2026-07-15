using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Get the first worksheet (default sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Add column headers for the sales data
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        // Sample sales data (Month -> Sales)
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        double[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };

        // Populate the worksheet with the data
        for (int i = 0; i < months.Length; i++)
        {
            // Row index starts at 2 because row 1 contains headers
            sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A (Month)
            sheet.Cells[i + 2, 1].PutValue(sales[i]);    // Column B (Sales)
        }

        // Save the workbook to a file in XLSX format
        workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);
    }
}
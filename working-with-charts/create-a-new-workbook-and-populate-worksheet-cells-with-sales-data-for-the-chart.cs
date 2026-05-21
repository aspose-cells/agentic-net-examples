using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header row for the sales data
        worksheet.Cells["A1"].PutValue("Month");
        worksheet.Cells["B1"].PutValue("Sales");

        // Sample sales data (Month and corresponding sales amount)
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        double[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };

        // Populate the worksheet with the sales data
        for (int i = 0; i < months.Length; i++)
        {
            // Column A (index 0) for month names
            worksheet.Cells[i + 2, 0].PutValue(months[i]);

            // Column B (index 1) for sales figures
            worksheet.Cells[i + 2, 1].PutValue(sales[i]);
        }

        // Save the workbook to a file in XLSX format
        workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);
    }
}
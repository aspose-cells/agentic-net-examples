// Title: Create a Workbook and add monthly sales data with Aspose.Cells for .NET (C#)
// Description: This C# example shows how to instantiate a Workbook, write a header row and monthly sales figures to the first worksheet, and save the file as SalesData.xlsx – ready to serve as a chart data source.
// Keywords: Aspose.Cells | C# workbook creation | populate Excel cells | sales data spreadsheet | SaveFormat.Xlsx | chart data source | Excel automation .NET | Aspose.Cells example
// Common Searches: Aspose.Cells write data to cells C# | Create Excel file with sales data using Aspose.Cells | How to save workbook as .xlsx in C# | Aspose.Cells sample for chart data | C# generate sales report Excel
// Developer Intent: Generate an .xlsx workbook containing month‑wise sales values.
// Use Cases: Build a sales report that feeds a line or column chart | Provide structured data for business‑intelligence dashboards | Automate monthly sales spreadsheet creation in a .NET application | Export sales figures to Excel for client delivery
// AI Prompts: Add code to format the header row (bold, background color) and auto‑size columns. | Insert a line chart that references the Month and Sales columns. | Allow the worksheet name to be specified via a parameter. | Export the workbook to a memory stream for web download. | Include total and average calculations below the data range.

using System;
using Aspose.Cells;

namespace AsposeCellsSalesDataExample
{
    // This C# example shows how to instantiate a Workbook, write a header row and monthly sales figures to the first worksheet, and save the file as SalesData.xlsx – ready to serve as a chart data source.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate header row
            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["B1"].PutValue("Sales");

            // Sample sales data (Month, Sales)
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            double[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };

            // Fill the data rows
            for (int i = 0; i < months.Length; i++)
            {
                // Row index starts at 2 (Excel rows are 1‑based, Cells["A2"] is the first data row)
                int row = i + 2;
                worksheet.Cells[$"A{row}"].PutValue(months[i]);
                worksheet.Cells[$"B{row}"].PutValue(sales[i]);
            }

            // Save the workbook to a file (uses the Workbook.Save(string, SaveFormat) rule)
            workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook with sales data created successfully.");
        }
    }
}

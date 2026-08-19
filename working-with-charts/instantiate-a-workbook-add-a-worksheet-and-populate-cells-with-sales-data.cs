// Title: C# Aspose.Cells – Create Workbook, Add “SalesData” Sheet, Populate Monthly Sales, Save as XLSX
// Description: Shows how to instantiate an Aspose.Cells Workbook in C#, rename the default worksheet to SalesData, add Month and Sales headers, fill rows with sample data, and save the file as SalesReport.xlsx.
// Keywords: Aspose.Cells | C# | Create Workbook | Add Worksheet | Populate Cells | Save XLSX | Sales data | Excel automation | Write to Excel | Aspose.Cells example
// Common Searches: Aspose.Cells create workbook C# | How to add worksheet and write data with Aspose.Cells | Save Excel file as XLSX using Aspose.Cells .NET | Write monthly sales data to Excel with Aspose | C# code to generate sales report Excel Aspose.Cells
// Developer Intent: Generate and save an Excel workbook containing a SalesData worksheet with month‑wise sales figures.
// Use Cases: Automated monthly sales reporting for finance teams | Preparing data source for Excel chart generation | Exporting sales figures from a .NET application to Excel for downstream BI tools | Creating template‑driven Excel files for client deliveries | Batch generation of sales reports across multiple periods
// AI Prompts: Add a line chart to the SalesData sheet that plots months versus sales. | Format the header row with bold text, white font, and a dark blue background. | Include a formula that calculates the total sales and place it below the data. | Convert the workbook to CSV while preserving the SalesData sheet. | Modify the code to read month and sales values from a JSON file instead of hard‑coded arrays.

using System;
using Aspose.Cells;

// Shows how to instantiate an Aspose.Cells Workbook in C#, rename the default worksheet to SalesData, add Month and Sales headers, fill rows with sample data, and save the file as SalesReport.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the default worksheet and give it a meaningful name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SalesData";

        // Add header row
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        // Sample sales data
        string[] months = { "January", "February", "March", "April", "May" };
        double[] sales = { 12000.5, 15000, 13000, 17000, 16000 };

        // Populate the worksheet with the data
        for (int i = 0; i < months.Length; i++)
        {
            // Column A (0-indexed) for month names
            sheet.Cells[i + 2, 0].PutValue(months[i]);
            // Column B (0-indexed) for sales figures
            sheet.Cells[i + 2, 1].PutValue(sales[i]);
        }

        // Save the workbook to a file in XLSX format
        workbook.Save("SalesReport.xlsx", SaveFormat.Xlsx);
    }
}

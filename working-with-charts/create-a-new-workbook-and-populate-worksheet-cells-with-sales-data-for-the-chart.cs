// Title: C# – Create a Workbook and Fill a Worksheet with Monthly Sales Data using Aspose.Cells
// Description: Demonstrates how to instantiate a Workbook, add a header row, populate month‑wise sales figures, and save the file as XLSX with Aspose.Cells for .NET – a prerequisite step for chart creation.
// Keywords: Aspose.Cells C# | create workbook .NET | populate worksheet cells | monthly sales data Excel | save XLSX file | Excel chart data source | Aspose.Cells example | write data to Excel programmatically
// Common Searches: how to add header row and data with Aspose.Cells | populate Excel worksheet with sales numbers C# | save workbook as XLSX using Aspose.Cells | prepare data for chart in Aspose.Cells | Aspose.Cells sample code for sales report
// Developer Intent: Generate an Excel workbook, write month and sales values into cells, and store the result as an XLSX file ready for charting.
// Use Cases: Build a basic sales report template before inserting a chart. | Provide a data source for line, column, or pie charts in Aspose.Cells. | Export program‑generated sales figures for downstream analysis or sharing.
// AI Prompts: Show me how to add a line chart that references the month and sales columns in this workbook. | Provide code to style the header row (bold, background color, alignment) using Aspose.Cells. | Explain how to make the data range dynamic so new months are automatically included in the chart.

using System;
using Aspose.Cells;

// Demonstrates how to instantiate a Workbook, add a header row, populate month‑wise sales figures, and save the file as XLSX with Aspose.Cells for .NET – a prerequisite step for chart creation.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Get the first worksheet (default sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row for the sales data
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        // Sample sales data (Month and corresponding sales amount)
        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        double[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };

        // Populate the worksheet with the sample data
        for (int i = 0; i < months.Length; i++)
        {
            // Column A (0-index) for month names
            sheet.Cells[i + 1, 0].PutValue(months[i]);

            // Column B (1-index) for sales figures
            sheet.Cells[i + 1, 1].PutValue(sales[i]);
        }

        // Save the workbook to an XLSX file
        workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);
    }
}

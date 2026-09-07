// Title: Create a column chart for monthly sales data in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that inserts a column chart into a worksheet, uses the range A1:B6 as the data source, and sets the chart title to "Monthly Sales". | Show how to position a column chart on a sheet, define its data range, and save the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# example for adding a column chart with a data range | how to set chart title in Aspose.Cells column chart | programmatically create sales column chart in Excel using Aspose.Cells .NET | Aspose.Cells set chart position and size for column chart
// Tags: Aspose.Cells add column chart | Aspose.Cells set chart data range | Aspose.Cells set chart title | Aspose.Cells save workbook as xlsx | Aspose.Cells define chart location

using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills cells A1:B6 with month and sales values, adds a column chart covering that range, sets the chart title to "Monthly Sales", positions the chart on the sheet, and saves the file as SalesChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with sample sales data
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
        int[] sales = { 1200, 1500, 1100, 1800, 1600 };

        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
            sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
        }

        // Add a column chart to the worksheet (using the Add method rule)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (vertical plotting)
        chart.SetChartDataRange("A1:B6", true);

        // Set a title for the chart
        chart.Title.Text = "Monthly Sales";

        // Save the workbook
        workbook.Save("SalesChart.xlsx", SaveFormat.Xlsx);
    }
}

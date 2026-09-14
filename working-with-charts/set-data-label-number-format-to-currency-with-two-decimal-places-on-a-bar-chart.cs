// Title: Format bar chart data labels as currency with two decimal places using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook that adds a bar chart and displays its series data labels in currency format ($#,##0.00) with Aspose.Cells for .NET. | Write C# code to create a bar chart, enable data labels, and apply a two‑decimal‑place currency number format to those labels using Aspose.Cells. | Create a .xlsx file containing sample monthly amounts and a bar chart whose data labels are formatted as currency via the "$#,##0.00" format string in Aspose.Cells.
// Common Searches: Aspose.Cells C# format chart data labels as currency with two decimal places | set number format for bar chart data labels Aspose.Cells .NET | C# Aspose.Cells example currency data labels on bar chart | how to apply $#,##0.00 format to chart series labels using Aspose.Cells | create bar chart with currency labels in Excel using Aspose.Cells for .NET
// Tags: Aspose.Cells chart data label currency format | C# set chart series data label number format | Aspose.Cells bar chart formatted data labels | Excel export bar chart currency labels .NET | Aspose.Cells number format string for chart labels

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, inserts sample category and amount data, adds a bar chart, enables data labels for the first series, sets the data label number format to "$#,##0.00" (currency with two decimal places), and saves the file as BarChartWithCurrencyLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the bar chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(1234.56);
            sheet.Cells["B3"].PutValue(7890.12);
            sheet.Cells["B4"].PutValue(3456.78);

            // Add a bar chart to the worksheet (rows 5-20, columns 0-7)
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Show data labels for the series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Set data label number format to currency with two decimal places
            chart.NSeries[0].DataLabels.NumberFormat = "$#,##0.00";

            // Determine output file path
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "BarChartWithCurrencyLabels.xlsx");

            // Save the workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to: {outputFile}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Save a Workbook with a Column Chart as XLSX using Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, populate it with sample data, add a column chart, and persist the entire file—including the chart—to an XLSX document (ChartWorkbook.xlsx) with proper exception handling.
// Keywords: Aspose.Cells save chart workbook | export chart to XLSX C# | Aspose.Cells column chart example | C# write Excel file with chart | Aspose.Cells .NET chart export
// Common Searches: how to save an Aspose.Cells workbook that contains a chart | C# Aspose.Cells create column chart and export to XLSX | save chart‑filled Excel file using Aspose.Cells for .NET | Aspose.Cells example saving workbook with embedded chart
// Developer Intent: Persist the generated workbook, which includes a column chart, as an XLSX file.
// Use Cases: Automated generation of sales dashboards that need to be shared as standard Excel files. | Scheduled creation of performance reports with visual charts for downstream business processes. | Building printable Excel workbooks that contain charts for monthly review meetings.
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart from a data range and saves the workbook as XLSX. | Show an example of robust error handling when saving an Aspose.Cells workbook containing multiple charts. | Provide code to add a title, axis labels, and legend to a chart before exporting the workbook to XLSX with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a new Workbook, populate it with sample data, add a column chart, and persist the entire file—including the chart—to an XLSX document (ChartWorkbook.xlsx) with proper exception handling.
    public class SaveChartWorkbookDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // Insert a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Define the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Save the workbook (including the chart) as an XLSX file
                workbook.Save("ChartWorkbook.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully as ChartWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SaveChartWorkbookDemo.Run();
        }
    }
}

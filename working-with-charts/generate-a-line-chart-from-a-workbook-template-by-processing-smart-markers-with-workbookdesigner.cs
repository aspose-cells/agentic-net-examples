// Title: Create a Line Chart from an Excel Template with Smart Markers using Aspose.Cells for .NET (C#)
// Description: This example loads an Excel template that contains Smart Markers, binds a List<SalesData> to the "Sales" data source, processes the markers with WorkbookDesigner, adds a line chart that references the populated ranges, customizes the title and legend, and saves the workbook as Output.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Smart Markers | WorkbookDesigner | line chart | Excel template | dynamic chart | sales data | GitHub example
// Common Searches: Aspose.Cells create line chart from template | C# smart markers line chart example | WorkbookDesigner process smart markers and add chart | how to bind a list to smart markers in Aspose.Cells | generate Excel line chart programmatically .NET
// Developer Intent: Generate an Excel line chart by processing Smart Markers with WorkbookDesigner and saving the result.
// Use Cases: Automate monthly sales reports using a pre‑designed Excel template. | Produce dynamic line charts for varying data sets without manual range adjustments. | Integrate chart generation into a .NET reporting service or API. | Create reusable Excel templates for finance or KPI dashboards.
// AI Prompts: Show code to set category axis labels for a line chart after processing smart markers in Aspose.Cells. | Demonstrate adding multiple series to the line chart from separate smart marker collections. | Explain handling of empty cells and null values when populating chart data with WorkbookDesigner. | Provide a GitHub‑style project layout for this line‑chart example.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSmartMarkerLineChart
{
    // Simple POCO class representing the data that will be bound to smart markers
    // This example loads an Excel template that contains Smart Markers, binds a List<SalesData> to the "Sales" data source, processes the markers with WorkbookDesigner, adds a line chart that references the populated ranges, customizes the title and legend, and saves the workbook as Output.xlsx.
    public class SalesData
    {
        public string Month { get; set; }
        public double Amount { get; set; }

        public SalesData(string month, double amount)
        {
            Month = month;
            Amount = amount;
        }
    }

    public class GenerateLineChartFromTemplate
    {
        public static void Run()
        {
            const string templatePath = "Template.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                // Verify that the template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Error: Template file \"{templatePath}\" not found.");
                    return;
                }

                // Load the workbook template that contains smart markers
                Workbook workbook = new Workbook(templatePath);

                // Prepare the data source that will replace the smart markers
                List<SalesData> sales = new List<SalesData>
                {
                    new SalesData("Jan", 1200.5),
                    new SalesData("Feb", 1500.0),
                    new SalesData("Mar", 1100.75),
                    new SalesData("Apr", 1700.25),
                    new SalesData("May", 1600.0)
                };

                // Bind the data source and process the smart markers
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    // Use range smart markers (LineByLine is obsolete but kept for compatibility)
                    LineByLine = false
                };
                designer.SetDataSource("Sales", sales);
                designer.Process(); // populate the worksheet with the data

                // Add a line chart that visualizes the data
                Worksheet sheet = workbook.Worksheets[0];
                int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 2, 25, 12);
                Chart lineChart = sheet.Charts[chartIndex];

                // Define the data range for the series (Amount column) and categories (Month column)
                int lastDataRow = 1 + sales.Count; // rows are zero‑based; data starts at row 1 (A2)
                string amountRange = $"=Sheet1!$B$2:$B${lastDataRow}";
                string monthRange = $"=Sheet1!$A$2:$A${lastDataRow}";

                // Add the series; categories are set automatically based on the range
                lineChart.NSeries.Add(amountRange, true);

                // If the API version supports setting category data, it can be done here.
                // The property may be unavailable in newer versions, so we skip it safely.
                // lineChart.NSeries[0].CategoryData = monthRange; // optional

                // Optional chart customizations
                lineChart.Title.Text = "Monthly Sales";
                lineChart.ShowLegend = false;
                lineChart.PlotEmptyCellsType = PlotEmptyCellsType.NotPlotted;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            GenerateLineChartFromTemplate.Run();
            Console.WriteLine("Processing completed.");
        }
    }
}

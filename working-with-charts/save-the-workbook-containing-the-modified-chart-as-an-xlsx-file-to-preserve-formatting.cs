// Title: Save a workbook containing a customized column chart as XLSX with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, adds a column chart, changes its title, and saves it as an XLSX file using Aspose.Cells. | Write a .NET example that modifies a chart's properties in an Aspose.Cells workbook and persists the workbook in XLSX format while keeping chart formatting. | Provide a snippet to export an Aspose.Cells workbook with a customized chart to an XLSX file, ensuring the chart appearance is retained.
// Common Searches: how to export an Aspose.Cells workbook with a modified chart to XLSX in C# | Aspose.Cells save chart formatting when saving workbook as .xlsx | C# example for adding a column chart and saving as XLSX using Aspose.Cells | preserve column chart title changes in Aspose.Cells XLSX output | Aspose.Cells .NET save workbook with updated chart properties
// Tags: save workbook as xlsx Aspose.Cells | modify chart title Aspose.Cells C# | export column chart to xlsx .NET | preserve chart formatting Aspose.Cells | add column chart Aspose.Cells example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveExample
{
    // The program creates a new workbook, fills cells with sample data, adds a column chart, sets the chart title to "Fruit Sales", and then saves the workbook as ModifiedChart.xlsx in XLSX format, preserving the chart's formatting.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Modify chart properties (e.g., title) to demonstrate a change
            chart.Title.Text = "Fruit Sales";

            // Save the workbook as XLSX to preserve formatting (lifecycle: save)
            workbook.Save("ModifiedChart.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook with modified chart saved as 'ModifiedChart.xlsx'.");
        }
    }
}

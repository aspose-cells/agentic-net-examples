// Title: Save a Workbook with a Column Chart to XLSX using Aspose.Cells for .NET
// Description: Creates a new workbook, inserts sample data, adds a column chart, and saves the workbook—including the chart—as an XLSX file with Workbook.Save and SaveFormat.Xlsx.
// Keywords: Aspose.Cells save workbook XLSX | C# Aspose.Cells chart export | Workbook.Save with chart | Aspose.Cells column chart example | export Excel chart .NET
// Common Searches: Aspose.Cells save workbook containing chart | C# save chart to XLSX with Aspose.Cells | Workbook.Save example for charts | How to export Aspose.Cells chart to XLSX | Save Excel file with chart using Aspose.Cells
// Developer Intent: Persist a workbook that includes a generated column chart as an XLSX file.
// Use Cases: Automated generation of sales dashboards with charts for distribution as XLSX files. | Web API that creates Excel reports with visualizations and returns them to clients. | Scheduled batch process that builds Excel workbooks with charts for archival storage.
// AI Prompts: Generate C# code that creates a line chart from a data range and saves the workbook as XLSX with Aspose.Cells. | Show how to add multiple data series to an Aspose.Cells chart before saving the workbook. | Explain how to set chart titles, axis labels, and legends, then export the workbook to XLSX using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveDemo
{
    // Creates a new workbook, inserts sample data, adds a column chart, and saves the workbook—including the chart—as an XLSX file with Workbook.Save and SaveFormat.Xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Fruits");
            worksheet.Cells["A3"].PutValue("Vegetables");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);          // Values
            chart.NSeries.CategoryData = "A2:A3";      // Categories

            // Save the workbook containing the chart as an XLSX file
            // Uses the Workbook.Save(string, SaveFormat) method rule
            workbook.Save("ChartWorkbook.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook with chart saved as ChartWorkbook.xlsx");
        }
    }
}

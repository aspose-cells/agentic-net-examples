// Title: Save Modified Chart Controls to XLSX with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, customize its title, and save the file as XLSX using Aspose.Cells Workbook.Save with SaveFormat.Xlsx so that all chart properties and controls are retained.
// Keywords: Aspose.Cells C# save chart | preserve chart controls Aspose.Cells | Workbook.Save SaveFormat.Xlsx | export Excel chart with title .NET | Aspose.Cells chart customization | C# Excel chart export | save modified chart Aspose.Cells | XLSX chart retention Aspose
// Common Searches: How to keep chart title when saving Aspose.Cells workbook to XLSX | Aspose.Cells preserve chart formatting on export | C# save Excel chart without losing settings | Aspose.Cells chart controls not lost after Save | Save modified chart to XLSX using Aspose.Cells
// Developer Intent: Export a workbook that contains edited chart elements while ensuring all chart settings remain intact.
// Use Cases: Automated sales reporting that adds a column chart, sets a custom title, and delivers an XLSX file with the chart fully rendered. | Dynamic dashboard generation where legends, data labels, or axis titles are programmatically changed and must survive the save operation. | Batch processing of Excel templates that require chart modifications before distribution to end‑users.
// AI Prompts: Generate C# code with Aspose.Cells to create a line chart, update its legend text, and save the workbook as XLSX preserving all chart controls. | Show an example that modifies multiple chart properties (title, axis labels, data labels) in Aspose.Cells and exports the file without losing any formatting. | Explain why Workbook.Save with SaveFormat.Xlsx retains chart objects and their settings in the resulting Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveDemo
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, customize its title, and save the file as XLSX using Aspose.Cells Workbook.Save with SaveFormat.Xlsx so that all chart properties and controls are retained.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Modify a chart control – for example, set a title
            chart.Title.Text = "Quarterly Sales";

            // Preserve all chart controls by saving the workbook in XLSX format
            // (using the documented Save(string, SaveFormat) method)
            workbook.Save("ModifiedCharts.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook with modified chart saved successfully.");
        }
    }
}

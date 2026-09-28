// Title: How to add the same .crtx chart template to every worksheet in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that loops through all worksheets in a Workbook and inserts a chart using a .crtx template, falling back to a default column chart when the template file is absent. | Demonstrate loading a chart template into a byte array and applying it to each worksheet while setting a common data range and chart title with Aspose.Cells for .NET.
// Common Searches: how to programmatically add a .crtx chart to every worksheet using Aspose.Cells in C# | loop over worksheets and create column charts with a shared data range in Aspose.Cells .NET | Aspose.Cells fallback to default chart when chart template file is missing | set chart title and position for multiple sheets in a workbook with Aspose.Cells
// Tags: apply .crtx chart template Aspose.Cells C# | add chart to each worksheet Aspose.Cells | iterate worksheets chart insertion Aspose.Cells .NET | set chart data range programmatically Aspose.Cells | fallback default column chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook with three worksheets, fills each with sample data, attempts to load a .crtx chart template, and then iterates over every worksheet to add a chart. If the template is available, it is applied; otherwise a default column chart is created. Each chart uses the same data range (A1:B6), receives a common title, and is positioned consistently before the workbook is saved as WorkbookWithCharts.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet to "Sheet1"
            Worksheet defaultSheet = workbook.Worksheets[0];
            defaultSheet.Name = "Sheet1";

            // Add additional worksheets
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Populate each worksheet with identical sample data
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                for (int i = 1; i <= 5; i++)
                {
                    ws.Cells[$"A{i + 1}"].PutValue($"Item {i}");
                    ws.Cells[$"B{i + 1}"].PutValue(i * 10);
                }
            }

            // Load a chart template (.crtx) into a byte array if the file exists
            byte[] templateData = null;
            const string templatePath = "ChartTemplate.crtx";
            if (File.Exists(templatePath))
            {
                templateData = File.ReadAllBytes(templatePath);
            }
            else
            {
                Console.WriteLine($"Template file '{templatePath}' not found. Charts will be created without a template.");
            }

            // Loop through each worksheet and add a chart (using template if available)
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Define the data range that the chart will use
                string dataRange = "A1:B6";

                int chartIdx;
                if (templateData != null)
                {
                    // Add a chart with the preset template
                    // Parameters: template bytes, data range, isVertical, topRow, leftColumn, bottomRow, rightColumn
                    chartIdx = ws.Charts.Add(templateData, dataRange, true, 5, 0, 20, 8);
                }
                else
                {
                    // Add a default chart without a template
                    chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    // Set data range; true indicates vertical orientation
                    ws.Charts[chartIdx].SetChartDataRange(dataRange, true);
                }

                Chart chart = ws.Charts[chartIdx];
                chart.Title.Text = "Template Chart";
            }

            // Save the workbook with the added charts
            workbook.Save("WorkbookWithCharts.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

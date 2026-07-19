// Title: Create a Chart Template and Apply It to Column Charts with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a workbook, fill a data range, add a column chart, apply a reusable built‑in chart style (Style 5) as a template, set a title, and save the file as ChartWithAppliedTemplate.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells chart template C# | apply chart style Aspose.Cells .NET | column chart built‑in style | save workbook with chart Aspose | C# Excel chart automation | Aspose.Cells chart formatting US | Aspose.Cells chart template Europe
// Common Searches: how to create a chart template in Aspose.Cells C# | apply built‑in chart style programmatically Aspose.Cells | C# add column chart and set style Aspose.Cells | Aspose.Cells reusable chart formatting example | generate styled Excel charts with Aspose.Cells .NET
// Developer Intent: Generate an Excel workbook that contains a column chart styled via a reusable template and persist it as an XLSX file.
// Use Cases: Standardize visual appearance across monthly sales dashboards. | Automate report generation where every workbook needs the same chart layout and style. | Batch‑process data sets and attach identically formatted charts for corporate presentations.
// AI Prompts: Show me how to create a reusable chart template file in Aspose.Cells and apply it to new charts using C#. | Provide C# code to export a styled Aspose.Cells chart as PNG after applying a built‑in template. | Explain how to copy chart formatting from one workbook to another with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplateDemo
{
    // Demonstrates how to build a workbook, fill a data range, add a column chart, apply a reusable built‑in chart style (Style 5) as a template, set a title, and save the file as ChartWithAppliedTemplate.xlsx using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // Step 1: Create a workbook and add a chart with desired styling
                // -----------------------------------------------------------------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Sample data for the chart
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["A3"].PutValue("B");
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["B2"].PutValue(15);
                ws.Cells["B3"].PutValue(25);
                ws.Cells["B4"].PutValue(35);

                // Add a column chart
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = ws.Charts[chartIdx];
                chart.SetChartDataRange("A1:B4", false);
                chart.Title.Text = "Chart Using Template Style";
                chart.Style = 5; // Apply a built‑in style for consistency

                // Save the workbook with the chart
                string outputPath = "ChartWithAppliedTemplate.xlsx";
                wb.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Batch add identical .crtx chart template to every worksheet in an Aspose.Cells workbook (C#)
// Description: Creates a workbook with multiple sheets, fills each with sample data, loads a .crtx chart template (if available), loops through all worksheets to insert the same chart (or a default column chart), sets a sheet‑specific title, and saves the file.
// Keywords: Aspose.Cells | C# chart template | .crtx | add chart to all worksheets | batch chart creation | loop worksheets | default column chart | dynamic chart title | Excel automation | chart template from byte array
// Common Searches: Aspose.Cells add same chart to every sheet | C# loop to insert .crtx chart in workbook | fallback to default chart when template missing Aspose.Cells | set chart title per worksheet Aspose.Cells | batch apply chart template to multiple worksheets
// Developer Intent: Insert an identical pre‑designed chart on each worksheet, using a .crtx template when it exists and a default column chart as a fallback.
// Use Cases: Generate a multi‑sheet sales report where each sheet shows the same pre‑styled chart. | Automate dashboard creation by adding a consistent chart to newly added worksheets in a data pipeline. | Provide a graceful fallback to a basic column chart when the .crtx template file is not found.
// AI Prompts: Write C# code with Aspose.Cells that loads a .crtx chart template and adds the chart to every worksheet, binding range A1:B5 and customizing the title with the sheet name. | Explain how to detect a missing .crtx file and create a default column chart instead, using Aspose.Cells APIs. | Show how to loop through all worksheets in a workbook and apply the same chart template efficiently. | Provide troubleshooting steps if the chart template is not applied correctly on some sheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartBatch
{
    // Creates a workbook with multiple sheets, fills each with sample data, loads a .crtx chart template (if available), loops through all worksheets to insert the same chart (or a default column chart), sets a sheet‑specific title, and saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and add a few worksheets
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Populate each worksheet with sample data that the chart will use
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Value");
                    for (int i = 2; i <= 5; i++)
                    {
                        ws.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                        ws.Cells[$"B{i}"].PutValue(i * 10);
                    }
                }

                // Load the chart template (.crtx) into a byte array if the file exists
                byte[] chartTemplateData = null;
                const string templatePath = "ChartTemplate.crtx";
                if (File.Exists(templatePath))
                {
                    chartTemplateData = File.ReadAllBytes(templatePath);
                }
                else
                {
                    Console.WriteLine($"Warning: Chart template '{templatePath}' not found. Charts will be created without a template.");
                }

                // Loop through each worksheet and add the same chart using the template (or default)
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    int chartIndex;
                    if (chartTemplateData != null)
                    {
                        // Add chart using the template
                        chartIndex = ws.Charts.Add(
                            chartTemplateData,   // template data
                            "A1:B5",             // data range
                            true,                // plot series by column (vertical)
                            5,                   // top row index
                            0,                   // left column index
                            15,                  // bottom row index
                            5);                  // right column index
                    }
                    else
                    {
                        // Add a default column chart and bind the data range manually
                        chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                        ws.Charts[chartIndex].NSeries.Add("A1:B5", true);
                    }

                    // Customize the chart title
                    Chart chart = ws.Charts[chartIndex];
                    chart.Title.Text = $"Sample Chart on {ws.Name}";
                }

                // Save the workbook with the added charts
                string outputPath = "WorkbookWithCharts.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

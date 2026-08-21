// Title: Aspose.Cells C# – Create a Chart from a .crtx Template and Apply Custom Data Labels
// Description: This example shows how to build a workbook, fill a simple data range, load a .crtx chart template (if it exists), add a column chart using the template or create one manually, and then customize the first series' data labels—showing values, setting the position to InsideEnd, applying a "0.00" number format, and styling the font (dark blue, size 12). The workbook is saved as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | chart template | .crtx | custom data labels | column chart | Excel workbook | sample code | GitHub example | API usage | data label font | number format
// Common Searches: load .crtx chart template Aspose.Cells C# | apply chart template to worksheet Aspose.Cells | set data label position and format Aspose.Cells | customize data label font Aspose.Cells column chart | fallback to manual chart creation when template missing Aspose.Cells
// Developer Intent: Load a .crtx chart template, add a chart to a worksheet, and configure custom data label properties.
// Use Cases: Read a ChartTemplate.crtx file into a byte array and create a chart with predefined styling. | Automatically switch to programmatic chart creation if the template file cannot be found. | Show series values on the chart, place labels inside the bar ends, apply a numeric format, and style the label font.
// AI Prompts: Generate C# code that creates a line chart from a .crtx template and sets data label font to red, size 10, with a custom number format. | Explain how to edit a .crtx chart template to embed default data label settings before using it with Aspose.Cells. | Provide a step‑by‑step guide for handling missing chart template files and falling back to manual chart creation in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExample
{
    // This example shows how to build a workbook, fill a simple data range, load a .crtx chart template (if it exists), add a column chart using the template or create one manually, and then customize the first series' data labels—showing values, setting the position to InsideEnd, applying a "0.00" number format, and styling the font (dark blue, size 12). The workbook is saved as an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Load a chart template (.crtx) into a byte array if the file exists
                byte[] templateData = null;
                const string templatePath = "ChartTemplate.crtx";

                if (File.Exists(templatePath))
                {
                    templateData = File.ReadAllBytes(templatePath);
                }
                else
                {
                    Console.WriteLine($"Template file \"{templatePath}\" not found. The chart will be created without a template.");
                }

                int chartIdx;

                if (templateData != null)
                {
                    // Add a chart using the template. The Add method with a byte[] parameter applies the preset template.
                    chartIdx = sheet.Charts.Add(
                        templateData,          // template byte array
                        "A1:B4",              // data range for the chart
                        true,                 // plot series by column
                        5,                    // top row of the chart
                        0,                    // left column of the chart
                        20,                   // bottom row of the chart
                        8);                   // right column of the chart
                }
                else
                {
                    // Create a chart without a template and set its data source manually
                    chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    Chart tempChart = sheet.Charts[chartIdx];
                    tempChart.NSeries.Add("A1:B4", true);
                }

                Chart chart = sheet.Charts[chartIdx];

                // Ensure the chart type matches the intended type (optional)
                chart.Type = ChartType.Column;

                // Access the first series and configure its data labels
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true; // show values
                series.DataLabels.Position = LabelPositionType.InsideEnd; // position inside the end of bars
                // Note: DataLabels does not expose a ShapeType property; omitted to avoid compilation error.
                series.DataLabels.NumberFormat = "0.00"; // custom number format
                series.DataLabels.Font.Color = Color.DarkBlue;
                series.DataLabels.Font.Size = 12;
                series.DataLabels.ApplyFont(); // apply font settings to all child labels

                // Save the workbook with the chart
                const string outputPath = "ChartWithTemplate.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

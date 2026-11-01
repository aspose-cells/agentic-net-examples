// Title: Add a new series to an existing line chart in an Excel workbook and export it as a PDF with a textbox annotation using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add additional series to an existing line chart and save as PDF | How to insert a new data series into a chart in an Excel file with Aspose.Cells .NET | Export updated Excel chart with textbox annotation to PDF using Aspose.Cells | C# code to add column E as a series to a line chart and generate PDF with Aspose.Cells | Add textbox note to chart after updating series in Aspose.Cells workbook
// Tags: add data series to line chart Aspose.Cells | export updated chart to PDF Aspose.Cells | textbox annotation on worksheet Aspose.Cells | use string range formula for chart series Aspose.Cells | programmatically modify Excel chart C# Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads an existing workbook (or creates a sample one), ensures a line chart exists, adds a new series from column E using a string range formula, inserts a textbox annotation, recalculates formulas, and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "Timeline.xlsx";
            const string outputFile = "TimelineUpdated.pdf";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            Worksheet sheet;

            if (File.Exists(inputFile))
            {
                workbook = new Workbook(inputFile);
                sheet = workbook.Worksheets[0];
            }
            else
            {
                workbook = new Workbook();
                sheet = workbook.Worksheets[0];

                // Populate sample data (A1:D5) for demonstration
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["D1"].PutValue("Series3");
                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[i - 1, 0].PutValue(DateTime.Today.AddDays(i - 2));
                    sheet.Cells[i - 1, 1].PutValue(i * 10);
                    sheet.Cells[i - 1, 2].PutValue(i * 15);
                    sheet.Cells[i - 1, 3].PutValue(i * 20);
                }

                // Add a simple line chart as a placeholder
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                // Use a string formula for the series range (B2:B5)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = sheet.Cells[0, 1].StringValue;
                chart.Title.Text = "Sample Chart";

                // Save the template for future runs
                workbook.Save(inputFile);
            }

            // Ensure there is at least one chart; create one if missing
            Chart timeline;
            if (sheet.Charts.Count > 0)
            {
                timeline = sheet.Charts[0];
            }
            else
            {
                int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                timeline = sheet.Charts[chartIdx];
                timeline.Title.Text = "Generated Chart";
            }

            // Use a supported chart type (Timeline is not available)
            if (timeline.Type != ChartType.Line)
            {
                timeline.Type = ChartType.Line;
            }

            // Determine data range (excluding header row)
            int firstDataRow = 1; // zero‑based index (row 2 in Excel)
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Column E (index 4) will hold the new series values
            int newSeriesColumnIndex = 4;

            // Ensure column E has values; if not, add placeholder values
            if (lastDataRow < firstDataRow)
            {
                lastDataRow = firstDataRow;
            }
            for (int r = firstDataRow; r <= lastDataRow; r++)
            {
                if (sheet.Cells[r, newSeriesColumnIndex].Value == null)
                {
                    sheet.Cells[r, newSeriesColumnIndex].PutValue((r + 1) * 5);
                }
            }

            // Add the new series to the chart using a string formula (e.g., E2:E5)
            int seriesCount = timeline.NSeries.Count;
            string newSeriesFormula = $"E{firstDataRow + 1}:E{lastDataRow + 1}";
            timeline.NSeries.Add(newSeriesFormula, true);
            timeline.NSeries[seriesCount].Name = sheet.Cells[0, newSeriesColumnIndex].StringValue;

            // Add a textbox annotation indicating the update
            Shape annotation = sheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "New series added",
                "Arial",
                12,
                false,
                false,
                100,   // left (points)
                100,   // top (points)
                200,   // width (points)
                30,    // height (points)
                0,     // text direction
                0);    // text alignment

            // Optional styling (commented out if not supported in current version)
            //annotation.Fill.ForeColor = Color.Yellow;
            //annotation.Line.ForeColor = Color.Black;

            // Recalculate formulas if any
            workbook.CalculateFormula();

            // Save the updated workbook as PDF
            workbook.Save(outputFile, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsHistogramDataLabels
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate source data ----------
                // Category labels
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Raw values for the histogram
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(25);
                sheet.Cells["B4"].PutValue(40);

                // Calculation range – for example, double the raw values
                sheet.Cells["C1"].PutValue("Calc");
                sheet.Cells["C2"].Formula = "=B2*2";
                sheet.Cells["C3"].Formula = "=B3*2";
                sheet.Cells["C4"].Formula = "=B4*2";

                // ---------- Add a column chart (used as histogram) ----------
                // ChartType.Column creates a simple column chart suitable for histogram‑style data
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data series (values) and the category (X‑axis) data
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ---------- Configure data labels ----------
                Series series = chart.NSeries[0];

                // Show the original value
                series.DataLabels.ShowValue = true;

                // Use the calculation range as the label text
                series.DataLabels.ShowCellRange = true;          // Enable cell‑range based labels
                series.DataLabels.LinkedSource = "C2:C4";        // Link to the calculated cells

                // Ensure the label shape automatically resizes to fit the text
                series.DataLabels.IsResizeShapeToFitText = true;

                // Optional: position the labels inside the bars
                series.DataLabels.Position = LabelPositionType.InsideBase;

                // Recalculate the chart so that all layout information is up‑to‑date
                chart.Calculate();

                // ---------- Save the workbook ----------
                string outputPath = "HistogramWithCalculatedDataLabels.xlsx";

                // Ensure the directory exists (handles cases where only a file name is provided)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
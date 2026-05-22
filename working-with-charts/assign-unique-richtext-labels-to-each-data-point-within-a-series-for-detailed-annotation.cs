using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRichTextDataLabels
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (categories in column A, values in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["A4"].PutValue("Gamma");
                sheet.Cells["B4"].PutValue(45);
                sheet.Cells["A5"].PutValue("Delta");
                sheet.Cells["B5"].PutValue(60);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Bind the data series to the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";      // Categories

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;        // Show the numeric value
                series.DataLabels.Position = LabelPositionType.OutsideEnd;

                // Iterate through each data point and assign a unique rich‑text label
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // Disable automatic text generation so we can set custom text
                    point.DataLabels.IsAutoText = false;

                    // Build a custom label (you can embed any rich‑text formatting you need)
                    // Example: "Item 1 – Alpha (15)"
                    string category = sheet.Cells[$"A{i + 2}"].StringValue;
                    double value = Convert.ToDouble(point.YValue); // Cast required for older API versions
                    string customText = $"Item {i + 1} – {category} ({value})";

                    // Assign the custom text to the data label
                    point.DataLabels.Text = customText;

                    // Optional: customize appearance of each label (font, color, etc.)
                    point.DataLabels.Font.Color = Color.DarkBlue;
                    point.DataLabels.Font.IsBold = true;
                    point.DataLabels.Font.Size = 10;
                }

                // Save the workbook to an XLSX file
                string outputPath = "RichTextDataLabels.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
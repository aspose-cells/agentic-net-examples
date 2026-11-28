// Title: How to apply a custom dashed line style to a line series in a Combo (Column‑Line) chart using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# set dash style for line series in combo chart | Create a combo chart with column and line series and custom line border in .NET | Change line series border to dashed dark red in Aspose.Cells Excel chart | Set marker size and colors for a line series in an Aspose.Cells combo chart | Use MsoLineDashStyle.Dash with Aspose.Cells to style chart lines in C#
// Tags: Aspose.Cells combo chart line series dash style | C# Aspose.Cells set chart series border dash | Aspose.Cells line series custom border color | Aspose.Cells chart marker size adjustment .NET | MsoLineDashStyle usage in Aspose.Cells charts

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ComboChartCustomLineStyle
{
    // The example creates a workbook, adds quarterly sales and profit data, inserts a Combo chart (column for sales, line for profit), and customizes the line series by applying a dashed border with medium weight and dark‑red color using MsoLineDashStyle.Dash. It also sets a circular marker, adjusts its size, and defines marker colors, then adds titles and saves the file as ComboChartWithDashedLine.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(45);
                sheet.Cells["C4"].PutValue(55);
                sheet.Cells["C5"].PutValue(70);

                // Add a Combo chart (Column + Line)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // First series – Column
                int colSeriesIdx = chart.NSeries.Add("B2:B5", true);
                Series colSeries = chart.NSeries[colSeriesIdx];
                colSeries.Name = "Sales";
                colSeries.Type = ChartType.Column;

                // Second series – Line
                int lineSeriesIdx = chart.NSeries.Add("C2:C5", true);
                Series lineSeries = chart.NSeries[lineSeriesIdx];
                lineSeries.Name = "Profit";
                lineSeries.Type = ChartType.Line;

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // Apply a custom dashed line style to the line series
                try
                {
                    // Use MsoLineDashStyle for dash patterns
                    lineSeries.Border.DashType = MsoLineDashStyle.Dash;
                    lineSeries.Border.Weight = WeightType.MediumLine;
                    lineSeries.Border.Color = Color.DarkRed;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to set line style – {ex.Message}");
                }

                // Optional: increase marker size for better visibility (if supported)
                try
                {
                    lineSeries.Marker.MarkerStyle = ChartMarkerType.Circle;
                    var sizeProp = lineSeries.Marker.GetType().GetProperty("Size");
                    if (sizeProp != null && sizeProp.CanWrite)
                    {
                        sizeProp.SetValue(lineSeries.Marker, 8, null);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to set marker size – {ex.Message}");
                }

                lineSeries.Marker.ForegroundColor = Color.DarkRed;
                lineSeries.Marker.BackgroundColor = Color.White;

                // Set chart title and axis titles
                chart.Title.Text = "Quarterly Sales and Profit";
                chart.Title.IsVisible = true;
                chart.CategoryAxis.Title.Text = "Quarter";
                chart.CategoryAxis.Title.IsVisible = true;
                chart.ValueAxis.Title.Text = "Amount";
                chart.ValueAxis.Title.IsVisible = true;

                // Save the workbook
                string outputPath = "ComboChartWithDashedLine.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

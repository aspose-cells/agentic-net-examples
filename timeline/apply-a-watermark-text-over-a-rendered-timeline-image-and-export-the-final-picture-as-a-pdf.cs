// Title: Add a Semi‑Transparent Text Watermark to a Timeline Chart Image and Export as PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells render timeline chart to PNG and add a text watermark before saving as PDF | How to apply a rotated semi‑transparent watermark to a PDF generated from an Excel chart with Aspose.Cells | Aspose.Cells RenderingWatermark example for adding CONFIDENTIAL watermark to chart image PDF in .NET | Export Excel timeline chart as high‑resolution PNG and embed in PDF with watermark using Aspose.Cells | Create PDF with watermark from chart image using Aspose.Cells C# code
// Tags: render timeline chart to PNG Aspose.Cells | apply RenderingWatermark to PDF Aspose.Cells | insert chart image into workbook Aspose.Cells | configure watermark opacity rotation Aspose.Cells | export workbook as PDF with watermark C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace WatermarkTimelinePdf
{
    // The sample loads or creates an Excel workbook containing a line chart used as a timeline, renders the chart to a high‑resolution PNG, inserts the image into a new workbook, defines a semi‑transparent rotated text watermark with RenderingWatermark, and saves the result as a PDF with the watermark applied, cleaning up the temporary image file afterwards.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths (adjust as needed)
            string sourceExcelPath = "timeline.xlsx";
            string tempImagePath = "timeline_image.png";
            string outputPdfPath = "TimelineWithWatermark.pdf";

            try
            {
                // ------------------------------------------------------------
                // 1. Load or create a workbook that contains a timeline chart
                // ------------------------------------------------------------
                Workbook sourceWorkbook;
                if (File.Exists(sourceExcelPath))
                {
                    sourceWorkbook = new Workbook(sourceExcelPath);
                }
                else
                {
                    // Create a simple workbook with a line chart that will act as a timeline
                    sourceWorkbook = new Workbook();
                    Worksheet ws = sourceWorkbook.Worksheets[0];
                    ws.Name = "Timeline";

                    // Populate sample data (Date, Value)
                    ws.Cells["A1"].PutValue("Date");
                    ws.Cells["B1"].PutValue("Value");
                    DateTime start = new DateTime(2023, 1, 1);
                    for (int i = 0; i < 10; i++)
                    {
                        ws.Cells[i + 2, 0].PutValue(start.AddDays(i));
                        ws.Cells[i + 2, 1].PutValue(100 + i * 10);
                    }

                    // Add a line chart (timeline)
                    int chartIdx = ws.Charts.Add(ChartType.Line, 5, 0, 20, 8);
                    Chart chart = ws.Charts[chartIdx];
                    chart.NSeries.Add("B2:B11", true);
                    chart.NSeries.CategoryData = "A2:A11";
                    chart.Title.Text = "Sample Timeline";

                    // Save the sample workbook for reference (optional)
                    sourceWorkbook.Save(sourceExcelPath);
                }

                // ------------------------------------------------------------
                // 2. Render the chart (timeline) to an image (PNG)
                // ------------------------------------------------------------
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
                if (sourceSheet.Charts.Count == 0)
                {
                    Console.WriteLine("No chart found in the source worksheet.");
                    return;
                }

                Chart timelineChart = sourceSheet.Charts[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    // Ensure the whole chart is captured
                    OnePagePerSheet = true
                };

                // Save chart image to a temporary file
                timelineChart.ToImage(tempImagePath, imgOptions);
                Console.WriteLine($"Chart rendered to image: {tempImagePath}");

                // ------------------------------------------------------------
                // 3. Create a new workbook and insert the rendered image
                // ------------------------------------------------------------
                Workbook finalWorkbook = new Workbook();
                Worksheet finalSheet = finalWorkbook.Worksheets[0];
                finalSheet.Name = "Result";

                if (!File.Exists(tempImagePath))
                {
                    Console.WriteLine("Rendered image file not found.");
                    return;
                }

                // Insert the image at the top‑left corner (cell A1)
                int picIdx = finalSheet.Pictures.Add(0, 0, tempImagePath);
                Picture pic = finalSheet.Pictures[picIdx];
                // No explicit size adjustment needed; Aspose.Cells sets picture size to image dimensions.

                // ------------------------------------------------------------
                // 4. Prepare a text watermark using RenderingWatermark
                // ------------------------------------------------------------
                RenderingFont watermarkFont = new RenderingFont("Arial", 48f)
                {
                    Bold = true,
                    Color = Color.FromArgb(150, Color.Red) // semi‑transparent red
                };

                RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
                {
                    // Center the watermark on the page
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    // Slight rotation for visual effect
                    Rotation = 30,
                    // 30 % opacity (0‑1 range)
                    Opacity = 0.3f,
                    // Place watermark in front of page contents
                    IsBackground = false,
                    // Scale to 50 % of page size
                    ScaleToPagePercent = 50
                };

                // ------------------------------------------------------------
                // 5. Save the workbook as PDF with the watermark applied
                // ------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                finalWorkbook.Save(outputPdfPath, pdfOptions);
                Console.WriteLine($"PDF generated with watermark: {outputPdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Clean up temporary image file
                if (File.Exists(tempImagePath))
                {
                    try { File.Delete(tempImagePath); } catch { /* ignore */ }
                }
            }
        }
    }
}

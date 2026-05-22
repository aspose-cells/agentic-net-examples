using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsComboChartPdf
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

                // Populate sample data for the combo chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                sheet.Cells["C1"].PutValue("Profit");
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(45);
                sheet.Cells["C4"].PutValue(60);
                sheet.Cells["C5"].PutValue(80);

                // Add a combo chart (Column + Line)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // First series – column
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Sales";

                // Second series – line
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Profit";
                chart.NSeries[1].Type = ChartType.Line;

                // NOTE: Older Aspose.Cells versions may not expose IsSecondaryValueAxis.
                // If needed, uncomment the line below when using a version that supports it.
                // chart.NSeries[1].IsSecondaryValueAxis = true;

                // Set category (X) axis labels
                chart.NSeries.CategoryData = "A2:A5";

                // Optional: give the chart a title
                chart.Title.Text = "Quarterly Sales and Profit";

                // Configure PDF save options to embed fonts
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedStandardWindowsFonts = true,
                    FontEncoding = PdfFontEncoding.Identity,
                    DefaultFont = "Arial",
                    CheckFontCompatibility = true
                };

                // Save the workbook (including the combo chart) as PDF
                string pdfPath = "ComboChart_EmbeddedFonts.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"Combo chart exported to PDF with embedded fonts: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
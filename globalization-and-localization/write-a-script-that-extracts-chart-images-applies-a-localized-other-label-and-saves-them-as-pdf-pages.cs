using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartExport
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains charts
            string sourcePath = "SourceWorkbook.xlsx";
            Workbook workbook = new Workbook(sourcePath);

            // Iterate through all worksheets
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];

                // Iterate through all charts in the current worksheet
                for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
                {
                    Chart chart = sheet.Charts[chartIndex];

                    // Apply a localized label – here we set the chart title to the word "Other"
                    // In a real scenario you could replace specific category names instead.
                    chart.Title.Text = "Other";

                    // ---------- Export chart as an image (PNG) ----------
                    string imageFile = $"Chart_{wsIndex}_{chartIndex}.png";
                    chart.ToImage(imageFile, ImageType.Png);

                    // ---------- Export chart as a PDF page ----------
                    // Define PDF page size (8.5 x 11 inches) and center alignment
                    string pdfFile = $"Chart_{wsIndex}_{chartIndex}.pdf";
                    chart.ToPdf(pdfFile, 8.5f, 11f,
                               PageLayoutAlignmentType.Center,
                               PageLayoutAlignmentType.Center);

                    Console.WriteLine($"Chart {chartIndex} on worksheet {wsIndex} exported to image and PDF.");
                }
            }

            // Optionally, save the modified workbook (titles changed) back to disk
            string modifiedWorkbook = "SourceWorkbook_Modified.xlsx";
            workbook.Save(modifiedWorkbook);
        }
    }
}
// Title: C# – Export Excel charts as PNG and centered 8.5×11 in PDF with localized “Other” label using Aspose.Cells
// Description: Loads an Excel workbook, replaces every cell containing the exact text "Other" with a supplied localized string (e.g., Chinese "其他"), then iterates through all worksheets and charts, exporting each chart to a PNG file and to a PDF page sized 8.5 × 11 inches with horizontal and vertical centering. The workbook is saved with the updated labels.
// Keywords: Aspose.Cells chart export PDF | C# export chart PNG | localize chart label Other | center chart on PDF page | batch chart extraction Aspose | Excel globalization Aspose.Cells | chart to image Aspose.Cells | replace cell value C#
// Common Searches: Aspose.Cells replace "Other" with localized text | export each Excel chart to separate PDF page C# | save chart as PNG and PDF with Aspose.Cells | center chart on 8.5x11 PDF using Aspose | globalize Excel chart labels programmatically
// Developer Intent: The developer needs to translate the "Other" category in an Excel workbook, then generate both PNG thumbnails and centered PDF pages for every chart, while preserving the localized workbook.
// Use Cases: Produce multilingual PDF reports where chart categories must be translated before distribution. | Create a printable catalog of all workbook charts with uniform 8.5 × 11 in pages centered for a professional look. | Generate PNG previews for web galleries while keeping the source workbook updated with localized labels.
// AI Prompts: Write a C# routine that scans a workbook, replaces cells equal to "Other" with a given localized string, and exports each chart to PNG and a centered 8.5×11 in PDF using Aspose.Cells. | Provide code to batch export all charts in an Excel file to both PNG and PDF formats, applying horizontal and vertical centering on the PDF pages. | Explain how to handle localization of chart category labels in Excel before exporting charts with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    // Loads an Excel workbook, replaces every cell containing the exact text "Other" with a supplied localized string (e.g., Chinese "其他"), then iterates through all worksheets and charts, exporting each chart to a PNG file and to a PDF page sized 8.5 × 11 inches with horizontal and vertical centering. The workbook is saved with the updated labels.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourceFile);

            // Localized label for "Other"
            string localizedOther = "其他";

            // Iterate through all worksheets
            for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
            {
                Worksheet sheet = workbook.Worksheets[wsIndex];

                // Replace any cell value equal to "Other" with the localized string
                // This ensures chart categories that use the cell value are updated
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.Type == CellValueType.IsString && cell.StringValue == "Other")
                    {
                        cell.PutValue(localizedOther);
                    }
                }

                // Process each chart in the worksheet
                for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
                {
                    Chart chart = sheet.Charts[chartIndex];

                    // Export chart image (PNG) – useful for further processing or verification
                    string imagePath = $"Chart_{wsIndex}_{chartIndex}.png";
                    chart.ToImage(imagePath, ImageType.Png);

                    // Export chart to a PDF page
                    // Desired page size: 8.5 x 11 inches, centered horizontally and vertically
                    string pdfPath = $"Chart_{wsIndex}_{chartIndex}.pdf";
                    chart.ToPdf(pdfPath, 8.5f, 11f,
                        PageLayoutAlignmentType.Center,
                        PageLayoutAlignmentType.Center);

                    Console.WriteLine($"Chart {chartIndex} on worksheet {wsIndex} exported to image and PDF.");
                }
            }

            // Optionally save the modified workbook (with localized labels)
            string modifiedWorkbookPath = "input_localized.xlsx";
            workbook.Save(modifiedWorkbookPath);
            Console.WriteLine($"Workbook saved with localized labels to '{modifiedWorkbookPath}'.");
        }
    }
}

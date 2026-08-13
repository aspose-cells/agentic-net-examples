// Title: Preserve three‑color scale conditional formatting when converting Excel to PDF with Aspose.Cells (.NET)
// Description: C# sample that builds a 11×11 multiplication table, applies a red‑yellow‑green three‑color scale conditional format to A1:K11, and saves the worksheet as a PDF using PdfSaveOptions with MergeAreas enabled and a light‑gray gridline color, so the color scale is retained in the PDF output.
// Keywords: Aspose.Cells | PDF export | conditional formatting | color scale | three color scale | MergeAreas | gridline color | .NET | C# | Excel to PDF | preserve colors | PdfSaveOptions | rendering options
// Common Searches: Aspose.Cells keep conditional formatting colors in PDF | How to export Excel with three‑color scale to PDF using .NET | PdfSaveOptions MergeAreas example Aspose.Cells | Set gridline color when saving workbook as PDF | Preserve color scale when converting workbook to PDF | C# Aspose.Cells PDF conditional formatting
// Developer Intent: Export an Excel workbook to PDF while retaining the visual colors of a three‑color scale conditional format.
// Use Cases: Generate printable reports where low, medium, and high values are highlighted with traffic‑light colors that must appear in the PDF. | Create PDF invoices or statements that preserve conditional formatting for discounts, taxes, or risk indicators. | Build PDF dashboards from Excel data that keep visual cues for quick data interpretation.
// AI Prompts: Show how to change the example to use a two‑color scale and still keep the colors in the exported PDF. | Explain the effect of the MergeAreas option on overlapping conditional‑formatting ranges during PDF rendering in Aspose.Cells. | Provide code that sets a custom gridline color, page margins, and header/footer while preserving conditional formatting in the PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsConditionalFormattingPdf
{
    // C# sample that builds a 11×11 multiplication table, applies a red‑yellow‑green three‑color scale conditional format to A1:K11, and saves the worksheet as a PDF using PdfSaveOptions with MergeAreas enabled and a light‑gray gridline color, so the color scale is retained in the PDF output.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (0 to 10) in a 11x11 range
                for (int i = 0; i <= 10; i++)
                {
                    for (int j = 0; j <= 10; j++)
                    {
                        sheet.Cells[i, j].PutValue(i * j);
                    }
                }

                // Add a three‑color scale conditional formatting to the range A1:K11
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

                // Define the area for the conditional formatting
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 10,
                    StartColumn = 0,
                    EndColumn = 10
                };
                fcc.AddArea(area);

                // Add the color scale condition
                int conditionIndex = fcc.AddCondition(FormatConditionType.ColorScale);
                FormatCondition colorScaleCondition = fcc[conditionIndex];
                colorScaleCondition.ColorScale.Is3ColorScale = true;
                colorScaleCondition.ColorScale.MinColor = Color.Red;      // Low values -> Red
                colorScaleCondition.ColorScale.MidColor = Color.Yellow;   // Mid values -> Yellow
                colorScaleCondition.ColorScale.MaxColor = Color.Green;    // High values -> Green

                // Configure PDF save options to preserve conditional formatting colors
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Merge conditional formatting areas before rendering so colors are kept
                    MergeAreas = true,

                    // Show gridlines with a light gray color
                    GridlineColor = Color.LightGray
                };

                // Save the workbook as PDF with the configured options
                workbook.Save("ConditionalFormattingColors.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

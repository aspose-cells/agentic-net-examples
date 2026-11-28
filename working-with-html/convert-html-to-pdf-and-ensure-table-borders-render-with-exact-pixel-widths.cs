// Title: Convert HTML to PDF in C# with exact 1‑pixel table borders using Aspose.Cells
// Common Searches: c# aspocells convert html to pdf preserving 1 pixel table borders | how to set thin cell borders for pixel‑perfect PDF output from html workbook | Aspose.Cells load html file into workbook and export as pdf with one page per sheet | adjust cell border line style to thin before saving html as pdf using Aspose.Cells | C# example for applying exact pixel border width in pdf generated from html
// Tags: Aspose.Cells HTML to PDF conversion | pixel‑perfect table borders in PDF | PdfSaveOptions OnePagePerSheet usage | LoadOptions for HTML workbook loading | adjust cell border weight in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfConverter
{
    // This C# console app validates input paths, loads an HTML file into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Html, iterates all cells to replace existing borders with the Thin line style (approximating a 1‑pixel border), and saves the workbook as a PDF with PdfSaveOptions configured for OnePagePerSheet. It also handles output directory creation and reports errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: HtmlToPdfConverter <inputHtmlPath> <outputPdfPath>");
                return;
            }

            string inputHtmlPath = args[0];
            string outputPdfPath = args[1];

            // Verify input HTML file exists
            if (!File.Exists(inputHtmlPath))
            {
                Console.WriteLine($"Error: Input HTML file not found at '{inputHtmlPath}'.");
                return;
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(outputPdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to create output directory:");
                Console.WriteLine(ex.Message);
                return;
            }

            try
            {
                // Load the HTML file into a Workbook.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

                // Adjust borders to approximate pixel widths (1 pixel ≈ 0.75 points).
                ApplyExactPixelBorderWidth(workbook, 0.75f);

                // Configure PDF save options.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                    // ImageResolution property is not available in the current Aspose.Cells version.
                };

                // Save the workbook as PDF.
                workbook.Save(outputPdfPath, pdfOptions);

                Console.WriteLine($"HTML successfully converted to PDF.\nInput : {inputHtmlPath}\nOutput: {outputPdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }

        /// <param name="workbook">The workbook containing the cells.</param>
        /// <param name="points">Border width in points (1 point = 1/72 inch).</param>
        private static void ApplyExactPixelBorderWidth(Workbook workbook, float points)
        {
            try
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Style style = cells[row, col].GetStyle();

                            bool hasBorder = false;
                            foreach (BorderType bt in Enum.GetValues(typeof(BorderType)))
                            {
                                if (style.Borders[bt].LineStyle != CellBorderType.None)
                                {
                                    hasBorder = true;
                                    style.Borders[bt].LineStyle = CellBorderType.Thin;
                                    // The Border class in this version does not expose a direct weight property.
                                    // Thin line style approximates a 1‑pixel border; further customization would require
                                    // a newer Aspose.Cells version that supports Border.Weight.
                                }
                            }

                            if (hasBorder)
                            {
                                cells[row, col].SetStyle(style);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while adjusting cell borders:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

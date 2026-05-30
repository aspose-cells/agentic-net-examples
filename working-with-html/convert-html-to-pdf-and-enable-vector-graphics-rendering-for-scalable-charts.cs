using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlToPdf
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source HTML file that contains the worksheet data and charts
                string htmlPath = "input.html";

                // Verify that the HTML file exists before attempting to load it
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: The file \"{htmlPath}\" was not found.");
                    return;
                }

                // Load the HTML file into a Workbook instance.
                // Aspose.Cells can directly load HTML documents.
                Workbook workbook = new Workbook(htmlPath);

                // ------------------------------------------------------------
                // OPTIONAL: Export each chart individually to PDF using the
                // Chart.ToPdf method. This method always renders charts as
                // vector graphics, ensuring they remain scalable.
                // ------------------------------------------------------------
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    for (int j = 0; j < sheet.Charts.Count; j++)
                    {
                        Chart chart = sheet.Charts[j];
                        string chartPdfPath = $"Chart_Sheet{i}_Chart{j}.pdf";

                        // Export the chart to PDF – vector rendering is guaranteed.
                        chart.ToPdf(chartPdfPath);
                        Console.WriteLine($"Chart exported to vector PDF: {chartPdfPath}");
                    }
                }

                // ------------------------------------------------------------
                // Save the entire workbook (including all charts) to a single PDF.
                // Charts are rendered as vector elements by default.
                // ------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Example: set PDF/A compliance (optional, does not affect vector rendering)
                    Compliance = PdfCompliance.PdfA1b
                };

                // The ImageType property is obsolete because charts are always rendered as vectors.
                // It is omitted here to avoid unnecessary dependencies.

                string pdfPath = "output.pdf";
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"Workbook saved to PDF with vector charts: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
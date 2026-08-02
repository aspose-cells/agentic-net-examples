// Title: C# – Convert CSV with Office Add‑In data to PDF while preserving clickable hyperlinks using Aspose.Cells
// Description: Creates a temporary CSV that includes text, hyperlinks, and numbers, then uses Aspose.Cells LoadOptions (LoadFormat.Csv) and PdfSaveOptions with ConversionUtility.Convert to generate a PDF that retains interactive elements such as clickable links. The sample also removes the temporary file after conversion.
// Keywords: Aspose.Cells CSV to PDF | clickable hyperlinks PDF conversion | Office Add‑In data export | C# Aspose.Cells ConversionUtility | PdfSaveOptions interactive elements | preserve interactivity .NET | CSV export to PDF example | Aspose.Cells LoadOptions CSV
// Common Searches: Aspose.Cells keep hyperlinks when converting CSV to PDF | C# convert Office Add‑In CSV to PDF with clickable links | ConversionUtility CSV to PDF sample code | PdfSaveOptions settings for interactive PDF in .NET | How to preserve interactivity during CSV‑to‑PDF conversion
// Developer Intent: Generate a PDF from a CSV produced by an Office Add‑In, ensuring that hyperlinks and other interactive controls remain functional.
// Use Cases: Produce PDF reports from CSV exports that contain URLs, allowing end‑users to click directly in the document. | Automate batch conversion of Office Add‑In CSV files to PDFs without losing interactivity. | Integrate CSV‑to‑PDF conversion into a .NET backend service that serves downloadable PDFs with active links.
// AI Prompts: Show how to set PDF page size and margins when converting a CSV to PDF with Aspose.Cells. | Demonstrate converting a CSV that includes embedded images to PDF while keeping the images clickable. | Explain how to add custom PDF metadata (title, author, keywords) during CSV‑to‑PDF conversion using ConversionUtility.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInConversion
{
    // Creates a temporary CSV that includes text, hyperlinks, and numbers, then uses Aspose.Cells LoadOptions (LoadFormat.Csv) and PdfSaveOptions with ConversionUtility.Convert to generate a PDF that retains interactive elements such as clickable links. The sample also removes the temporary file after conversion.
    class Program
    {
        static void Main()
        {
            // Paths for the temporary CSV file and the final PDF output
            string csvPath = "sample.csv";
            string pdfPath = "output.pdf";

            // Create a simple CSV file that could contain Office Add‑In data (e.g., formulas, hyperlinks)
            File.WriteAllText(csvPath,
                "Name,Link,Value\n" +
                "Item1,https://example.com,100\n" +
                "Item2,https://contoso.com,200");

            try
            {
                // LoadOptions tell Aspose.Cells to treat the source as a CSV file
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                // SaveOptions for PDF output – PdfSaveOptions derives from SaveOptions
                // (PdfSaveOptions is part of Aspose.Cells and does not require additional code)
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // Convert CSV directly to PDF while preserving any interactive elements
                // (e.g., hyperlinks become clickable in the PDF)
                ConversionUtility.Convert(csvPath, loadOptions, pdfPath, saveOptions);

                Console.WriteLine($"CSV file successfully converted to PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary CSV file
                if (File.Exists(csvPath))
                {
                    File.Delete(csvPath);
                }
            }
        }
    }
}

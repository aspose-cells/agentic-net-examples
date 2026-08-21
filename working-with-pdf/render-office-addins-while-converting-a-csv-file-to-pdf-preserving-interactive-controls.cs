// Title: Convert CSV to PDF in C# with Aspose.Cells ConversionUtility
// Description: Demonstrates how to create a temporary CSV file, validate its presence, load it using LoadOptions (CSV format), convert it to a PDF with PdfSaveOptions via Aspose.Cells ConversionUtility, and clean up the source file—all in a single C# console program.
// Keywords: Aspose.Cells | C# CSV to PDF | ConversionUtility | PdfSaveOptions | LoadOptions CSV | CSV to PDF conversion example | temporary file cleanup | Aspose.Cells API | office add‑ins rendering | interactive PDF controls
// Common Searches: Aspose.Cells convert CSV to PDF C# | C# code sample for CSV to PDF using Aspose | How to use ConversionUtility with CSV input | LoadOptions CSV Aspose.Cells example | PdfSaveOptions usage in Aspose.Cells | Delete temporary CSV after conversion C#
// Developer Intent: The developer needs a quick, reliable way to transform CSV data into a PDF document using Aspose.Cells in a C# application.
// Use Cases: Generate printable PDF reports directly from exported CSV data. | Automate nightly batch conversion of CSV logs to PDF archives while removing the original files. | Create PDF attachments for email campaigns from dynamically generated CSV content.
// AI Prompts: Show C# code that adds custom page margins and orientation when converting CSV to PDF with Aspose.Cells. | Explain how to embed clickable hyperlinks or form fields in the PDF generated from a CSV file. | Provide performance tips for converting large CSV files (100k+ rows) to PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

namespace CsvToPdfWithInteractiveControls
{
    // Demonstrates how to create a temporary CSV file, validate its presence, load it using LoadOptions (CSV format), convert it to a PDF with PdfSaveOptions via Aspose.Cells ConversionUtility, and clean up the source file—all in a single C# console program.
    class Program
    {
        static void Main()
        {
            // Paths for the source CSV and the resulting PDF
            string csvPath = "sample.csv";
            string pdfPath = "output.pdf";

            // Create a simple CSV file for demonstration
            File.WriteAllText(csvPath,
                "Name,Age,Score\n" +
                "Alice,30,85\n" +
                "Bob,25,92\n" +
                "Charlie,28,78");

            try
            {
                // Verify that the CSV file exists before attempting conversion
                if (!File.Exists(csvPath))
                {
                    throw new FileNotFoundException("The source CSV file was not found.", csvPath);
                }

                // LoadOptions specify that the source file is a CSV
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                // SaveOptions for PDF (no special properties needed for this conversion)
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // Convert CSV directly to PDF using the utility method
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
                    try
                    {
                        File.Delete(csvPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to delete temporary CSV file: {ex.Message}");
                    }
                }
            }
        }
    }
}

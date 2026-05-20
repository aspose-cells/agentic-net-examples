using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class ValidateTitleBeforeExport
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook instance
                Workbook workbook = new Workbook();

                // Access the built‑in document properties collection
                BuiltInDocumentPropertyCollection props = workbook.BuiltInDocumentProperties;

                // Example: set the Title property (comment out to test validation failure)
                // props.Title = "Sample Document Title";

                // Validate that the Title property is not null or empty before exporting
                string title = props.Title;
                if (string.IsNullOrWhiteSpace(title))
                {
                    throw new InvalidOperationException("The workbook's Title property must be set before exporting.");
                }

                // Export the workbook to PDF
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Display the document title in the PDF window title bar
                    DisplayDocTitle = true
                };

                // Save the workbook
                workbook.Save("ValidatedTitleExport.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateTitleBeforeExport.Run();
        }
    }
}
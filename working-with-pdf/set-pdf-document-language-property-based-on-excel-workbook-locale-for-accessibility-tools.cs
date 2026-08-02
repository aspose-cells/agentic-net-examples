// Title: Set PDF Language Metadata from Excel Workbook Locale using Aspose.Cells for .NET
// Description: Shows how to obtain an Excel workbook’s locale (Settings.LanguageCode or CultureInfo), convert it to a BCP‑47 tag, assign the built‑in Document Property "Language", and save the workbook as a PDF with Aspose.Cells so accessibility tools detect the correct language.
// Keywords: Aspose.Cells PDF language tag | embed workbook locale in PDF metadata | C# Aspose.Cells accessibility | Excel workbook language to PDF | BCP-47 language code | PDF document property Language | Aspose.Cells PDF metadata | locale‑aware PDF export | .NET PDF accessibility
// Common Searches: set language tag in PDF generated from Excel with Aspose.Cells | Aspose.Cells PDF language property based on workbook culture | C# embed accessibility language metadata when saving Excel to PDF | map workbook CountryCode to PDF language using Aspose.Cells | how to preserve Excel locale in PDF metadata
// Developer Intent: Add the correct language metadata to a PDF created from an Excel workbook, using the workbook’s locale for accessibility compliance.
// Use Cases: Create multilingual reports where screen readers need the proper language identifier. | Export regional Excel files to PDF while retaining the UI language for regulatory compliance. | Batch‑process Excel workbooks to PDF, automatically embedding the appropriate language tag.
// AI Prompts: Generate C# code with Aspose.Cells that reads a workbook’s CultureInfo and sets the PDF Language property before saving. | Convert an Aspose.Cells CountryCode enum value to a BCP‑47 language tag for PDF metadata in C#. | Explain how the built‑in Document Property "Language" is exported to PDF and how assistive technologies use it.

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;   // for PdfSaveOptions

namespace AsposeCellsPdfLanguageDemo
{
    // Shows how to obtain an Excel workbook’s locale (Settings.LanguageCode or CultureInfo), convert it to a BCP‑47 tag, assign the built‑in Document Property "Language", and save the workbook as a PDF with Aspose.Cells so accessibility tools detect the correct language.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: set the workbook's UI language (locale) to United States English
            // This influences culture‑specific formatting inside the workbook
            workbook.Settings.LanguageCode = CountryCode.USA;

            // Determine the language tag (e.g., "en-US") from the workbook's culture info.
            // If CultureInfo is not set, fall back to a default language tag.
            string languageTag = "en-US"; // default
            if (workbook.Settings.CultureInfo != null)
            {
                languageTag = workbook.Settings.CultureInfo.Name; // e.g., "en-US"
            }
            else
            {
                // Attempt to map CountryCode to a culture name
                try
                {
                    // CountryCode enum values match ISO country codes; prepend "en-" for English as a simple example
                    languageTag = $"en-{workbook.Settings.LanguageCode}";
                }
                catch { /* ignore mapping errors */ }
            }

            // Set the built‑in document property "Language" – this property is exported to PDF
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;
            builtInProps.Language = languageTag;

            // (Optional) Add some content to demonstrate the PDF
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample text for PDF language property.");

            // Configure PDF save options if needed (e.g., keep default font checking)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the document respects the language property for accessibility tools
                // No explicit property for language; the built‑in property is sufficient.
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as PDF; the Language property will be embedded in the PDF metadata
            workbook.Save("Output.pdf", pdfOptions);

            Console.WriteLine($"Workbook saved as PDF with Language property set to \"{languageTag}\".");
        }
    }
}

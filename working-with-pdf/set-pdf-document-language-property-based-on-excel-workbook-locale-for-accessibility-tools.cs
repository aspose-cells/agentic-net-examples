// Title: Set PDF Language Metadata from Excel Workbook Locale with Aspose.Cells for .NET
// Description: Demonstrates how to read a workbook's Settings.LanguageCode, map it to an IETF locale tag, assign the tag to the built‑in Document Property "Language", configure PdfSaveOptions.DefaultEditLanguage, and save the workbook as a PDF with proper language metadata for screen readers and other accessibility tools.
// Keywords: Aspose.Cells PDF language metadata | C# set PDF language tag | Workbook Settings.LanguageCode | PdfSaveOptions DefaultEditLanguage | built‑in document property Language | Excel locale to PDF | PDF accessibility Aspose.Cells | German PDF Aspose.Cells | CJK PDF Aspose.Cells
// Common Searches: Aspose.Cells set PDF language tag | Map Excel workbook locale to PDF language property | DefaultEditLanguage CJK Aspose.Cells | Set built‑in document property Language before PDF export | PDF accessibility language metadata .NET Aspose
// Developer Intent: Add language metadata to a PDF generated from an Excel workbook based on the workbook’s locale to support accessibility tools.
// Use Cases: Generate a German PDF (de‑DE) from a workbook whose LanguageCode is set to Germany, ensuring screen readers recognize the correct language. | Create PDFs for Japanese or Chinese markets where the workbook locale triggers DefaultEditLanguage.CJK for proper text editing and rendering. | Provide a reliable fallback to English (en‑US) when the workbook locale is not explicitly mapped, guaranteeing a valid Language property in every exported PDF.
// AI Prompts: Write C# code using Aspose.Cells that reads a workbook's CountryCode, maps it to an IETF locale tag, sets the built‑in Document Property "Language", and saves the file as PDF with appropriate DefaultEditLanguage. | Explain how PdfSaveOptions.DefaultEditLanguage influences PDF accessibility and how to configure it for Western versus CJK locales in Aspose.Cells. | Show how to extend the locale‑to‑tag mapping to include additional languages such as Spanish (es‑ES) and Italian (it‑IT) while preserving existing functionality.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to read a workbook's Settings.LanguageCode, map it to an IETF locale tag, assign the tag to the built‑in Document Property "Language", configure PdfSaveOptions.DefaultEditLanguage, and save the workbook as a PDF with proper language metadata for screen readers and other accessibility tools.
    public class PdfLanguageFromWorkbookLocale
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Set the workbook's UI language (locale) to German (Germany)
            // This influences the language used for accessibility tools.
            workbook.Settings.LanguageCode = CountryCode.Germany;

            // Retrieve the locale as a string in the format "de-DE"
            // Map a few common CountryCode values to locale tags.
            string localeTag = workbook.Settings.LanguageCode switch
            {
                CountryCode.USA => "en-US",
                CountryCode.UnitedKingdom => "en-GB",
                CountryCode.Germany => "de-DE",
                CountryCode.France => "fr-FR",
                CountryCode.Japan => "ja-JP",
                CountryCode.China => "zh-CN",
                _ => "en-US" // fallback
            };

            // Set the built‑in document property "Language" – this is exported to PDF.
            workbook.BuiltInDocumentProperties.Language = localeTag;

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the default edit language based on the locale.
                // Use CJK for East Asian locales, otherwise English.
                DefaultEditLanguage = localeTag switch
                {
                    "ja-JP" or "zh-CN" => DefaultEditLanguage.CJK,
                    _ => DefaultEditLanguage.English
                },

                // Ensure Unicode characters are rendered using the workbook's default font.
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as PDF with the language metadata applied.
            workbook.Save("Output.pdf", pdfOptions);

            Console.WriteLine($"Workbook saved as PDF with language tag '{localeTag}'.");
        }
    }
}

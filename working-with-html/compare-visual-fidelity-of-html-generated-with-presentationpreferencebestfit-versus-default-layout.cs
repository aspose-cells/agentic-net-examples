// Title: Aspose.Cells .NET: Compare HTML output with PresentationPreference (BestFit) vs default
// Description: Creates a workbook, then saves it twice as HTML – once with HtmlSaveOptions.PresentationPreference set to true (BestFit) and once with the default false setting – to let developers visually compare layout fidelity.
// Keywords: Aspose.Cells | PresentationPreference | BestFit | HTML export | .NET | C# | HtmlSaveOptions | visual fidelity | layout comparison | column width | text wrapping
// Common Searches: Aspose.Cells PresentationPreference true example | HTML export best fit Aspose.Cells | difference between PresentationPreference true and false | preserve column width in Aspose.Cells HTML | compare Aspose.Cells HTML layout
// Developer Intent: Generate two HTML files from the same workbook to evaluate how PresentationPreference (BestFit) changes the visual layout compared to the default export.
// Use Cases: Side‑by‑side visual comparison for reporting or documentation | Regression testing of HTML export appearance across Aspose.Cells versions | Ensuring column width and text‑wrapping consistency before publishing HTML | Creating HTML snapshots for UI review or client demos
// AI Prompts: Write C# code that loads Workbook_BestFit.html and Workbook_Default.html and highlights differences in column widths and text wrapping. | Explain how PresentationPreference modifies the generated CSS in Aspose.Cells HTML output and suggest additional HtmlSaveOptions to improve visual fidelity. | Provide a PowerShell script that opens both HTML files in the default browser for manual side‑by‑side comparison.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlComparison
{
    // Creates a workbook, then saves it twice as HTML – once with HtmlSaveOptions.PresentationPreference set to true (BestFit) and once with the default false setting – to let developers visually compare layout fidelity.
    class Program
    {
        static void Main()
        {
            // Create a sample workbook with test data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Presentation Preference Comparison");
            sheet.Cells["A2"].PutValue("Date:");
            sheet.Cells["B2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue("Value:");
            sheet.Cells["B3"].PutValue(12345.67);
            sheet.Cells["A4"].PutValue("Long Text:");
            sheet.Cells["B4"].PutValue("This is a long piece of text that will be used to demonstrate how the HTML output differs when PresentationPreference is enabled versus the default layout.");

            // -------------------------------------------------
            // Save HTML with PresentationPreference = true (BestFit)
            // -------------------------------------------------
            HtmlSaveOptions bestFitOptions = new HtmlSaveOptions();
            bestFitOptions.PresentationPreference = true;   // Enable presentation preference for better visual fidelity
            bestFitOptions.IsFullPathLink = false;          // Use relative links
            // Optional: keep default LayoutMode (Normal) for a fair comparison
            using (MemoryStream bestFitStream = new MemoryStream())
            {
                workbook.Save(bestFitStream, bestFitOptions);
                // Write the stream to a file for inspection
                File.WriteAllBytes("Workbook_BestFit.html", bestFitStream.ToArray());
                Console.WriteLine("HTML saved with PresentationPreference = true (BestFit).");
            }

            // -------------------------------------------------
            // Save HTML with default settings (PresentationPreference = false)
            // -------------------------------------------------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions();
            // PresentationPreference remains false (default)
            defaultOptions.IsFullPathLink = false;
            using (MemoryStream defaultStream = new MemoryStream())
            {
                workbook.Save(defaultStream, defaultOptions);
                File.WriteAllBytes("Workbook_Default.html", defaultStream.ToArray());
                Console.WriteLine("HTML saved with default PresentationPreference (false).");
            }

            // Inform the user that the comparison files are ready
            Console.WriteLine("Comparison files generated:");
            Console.WriteLine(" - Workbook_BestFit.html");
            Console.WriteLine(" - Workbook_Default.html");
        }
    }
}

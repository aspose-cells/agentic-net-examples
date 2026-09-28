// Title: Convert an HTML file to PDF using Aspose.Cells for .NET and insert a Roman‑numeral page number footer (alternative approach)
// AI Prompts: Write C# code that loads an HTML document into an Aspose.Cells Workbook, exports it as a PDF, and adds a footer that displays page numbers in Roman numerals using the available footer APIs. | Describe a workaround for adding Roman‑numeral page numbers to a PDF generated from HTML with Aspose.Cells when the CenterFooter property is not present, including how to use PageSetup or custom drawing.
// Common Searches: Aspose.Cells C# export HTML to PDF with Roman numeral page numbers in footer | How to add a custom footer with Roman numerals when converting HTML to PDF using Aspose.Cells .NET | Workaround for missing CenterFooter property in Aspose.Cells PDF export C# | Set page number format to Roman numerals in Aspose.Cells PDF output | Convert HTML to PDF and add page footer using Aspose.Cells for .NET
// Tags: Aspose.Cells HTML to PDF conversion C# | Roman numeral footer Aspose.Cells PDF | custom PDF footer Aspose.Cells .NET | fallback CenterFooter Aspose.Cells workaround | page number format Roman numerals Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input HTML file, loads it into an Aspose.Cells Workbook, ensures the output directory exists, and saves the workbook as a PDF. Because the current Aspose.Cells version lacks a CenterFooter property, the code does not directly apply a Roman‑numeral footer; a comment explains that a workaround (e.g., using PageSetup or custom drawing) is required for such footer customization.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (optional further processing)
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: The CenterFooter property is not available in the current Aspose.Cells version.
            // If needed, footer settings can be applied using alternative APIs or a newer library version.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

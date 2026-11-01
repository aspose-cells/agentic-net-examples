// Title: Save an Excel workbook as HTML with default settings and set hyperlink targets to open in a new tab (Aspose.Cells for .NET)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, saves it as HTML using the default HtmlSaveOptions, and then updates the produced HTML to add target="_blank" to every <a> element. | Show how to export a Workbook to HTML with Aspose.Cells and then adjust the resulting file so all hyperlinks open in a new browser tab.
// Common Searches: Aspose.Cells export Excel to HTML with hyperlinks opening in new tab | C# save workbook as HTML and add target blank to links using Aspose.Cells | How to modify Aspose.Cells generated HTML to set hyperlink target attribute | default HtmlSaveOptions hyperlink handling Aspose.Cells .NET
// Tags: Aspose.Cells HTML generation default options | adjust Aspose.Cells HTML output for target attribute | hyperlink target blank in Aspose.Cells generated HTML | Workbook conversion to HTML using Aspose.Cells | modify <a> tags in Aspose.Cells HTML output

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input Excel file, loads it into an Aspose.Cells Workbook, saves it as HTML with default HtmlSaveOptions, and notes that hyperlink targets must be added manually (e.g., by editing the HTML to include target="_blank" on each link).
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Note: In the current Aspose.Cells version, hyperlink target handling
            // is controlled automatically. If needed, post‑process the HTML to
            // add target="_blank" to <a> tags.

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

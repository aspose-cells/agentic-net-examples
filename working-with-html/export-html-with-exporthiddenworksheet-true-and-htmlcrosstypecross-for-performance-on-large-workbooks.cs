// Title: Export hidden worksheets to HTML with Aspose.Cells for .NET using HtmlSaveOptions for large workbooks
// Common Searches: Aspose.Cells C# export hidden sheets to HTML with HtmlSaveOptions | How to include hidden worksheets when saving Excel as HTML in .NET | Improve performance of HTML export for large Excel workbooks using Aspose.Cells | Set ExportHiddenWorksheet true in Aspose.Cells HTML conversion | Aspose.Cells HTML export options for large workbooks C#
// Tags: HtmlSaveOptions ExportHiddenWorksheet | Aspose.Cells HTML export hidden worksheets | large workbook HTML conversion Aspose.Cells | C# Aspose.Cells save workbook as HTML | performance optimization HTML export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, configures HtmlSaveOptions with ExportHiddenWorksheet set to true to include hidden sheets in the generated HTML, saves the workbook as an HTML file, and handles any runtime errors. It notes that HtmlCrossType is not available in the current Aspose.Cells version, making this approach suitable for large workbooks.
class ExportHtmlExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export hidden worksheets as part of the HTML output
                ExportHiddenWorksheet = true
                // Note: HtmlCrossType property is not available in the current Aspose.Cells version
            };

            // Save the workbook as HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

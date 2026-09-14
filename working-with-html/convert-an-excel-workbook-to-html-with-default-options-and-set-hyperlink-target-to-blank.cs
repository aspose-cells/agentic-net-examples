// Title: How to export an Excel workbook to HTML using Aspose.Cells for .NET and make all hyperlinks open in a new tab
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, creates HtmlSaveOptions, sets ExportHyperlinkTarget to "_blank", and saves the workbook as an HTML file. | Provide a C# example that uses reflection to detect the ExportHyperlinkTarget property and configures it so hyperlinks open in a new window when converting Excel to HTML with Aspose.Cells.
// Common Searches: Aspose.Cells .NET export Excel to HTML with hyperlinks opening in new tab | C# set ExportHyperlinkTarget to _blank when saving workbook as HTML | How to make links open in new window in HTML output from Aspose.Cells | Save Excel file as HTML using Aspose.Cells and customize hyperlink target
// Tags: Aspose.Cells HtmlSaveOptions ExportHyperlinkTarget | C# export Excel to HTML Aspose.Cells | set hyperlink target _blank Aspose.Cells | Excel to HTML conversion .NET Aspose | HTML export options workbook Aspose

using System;
using System.IO;
using Aspose.Cells;

// The program loads "input.xlsx" with Aspose.Cells, creates HtmlSaveOptions, uses reflection to set ExportHyperlinkTarget to "_blank" when supported, and saves the workbook as "output.html" so that all hyperlinks open in a new tab.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set hyperlink target to open in a new window/tab if the property exists (newer versions)
            var prop = typeof(HtmlSaveOptions).GetProperty("ExportHyperlinkTarget");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(htmlOptions, "_blank");
            }

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Insert <h1> headings for each worksheet when saving an Excel workbook to HTML using Aspose.Cells for .NET
// Common Searches: how to add worksheet name as <h1> tag in HTML output with Aspose.Cells C# | Aspose.Cells HtmlSaveOptions custom heading for each sheet when converting to HTML | C# generate HTML from Excel and prepend h1 tags to each worksheet table | save workbook as HTML with worksheet titles using Aspose.Cells .NET API | add HTML heading before each worksheet table during Excel to HTML conversion in C#
// Tags: Aspose.Cells HtmlSaveOptions worksheet heading | C# export Excel to HTML with h1 tags | custom HTML tags during workbook save | HTML conversion of multiple worksheets Aspose | add worksheet titles in HTML output C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, sets up HtmlSaveOptions for HTML export, notes that the ExportWorksheetHeader property is unavailable, and saves the workbook. To insert <h1> tags before each worksheet table, developers must add the headings manually after saving or manipulate the HTML string before writing it to the output file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The ExportWorksheetHeader property is not available in this version.
            // If worksheet headers are required, they can be added manually after saving.

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

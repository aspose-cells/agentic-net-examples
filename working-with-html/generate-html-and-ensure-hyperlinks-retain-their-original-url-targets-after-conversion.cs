// Title: Convert an Excel .xlsx workbook to HTML while keeping hyperlink URLs intact using Aspose.Cells for .NET
// AI Prompts: Generate HTML from a .xlsx file with Aspose.Cells, ensuring that each cell hyperlink retains its original URL. | Save a workbook as HTML in C# with HtmlSaveOptions so that all embedded hyperlinks are preserved.
// Common Searches: Aspose.Cells C# export workbook to HTML preserving hyperlink targets | how to keep Excel cell links when saving as HTML with Aspose.Cells | C# convert xlsx to html retain hyperlink URLs Aspose | HtmlSaveOptions hyperlink preservation Aspose.Cells example | save Excel file as HTML with active links using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions hyperlink export | C# .xlsx to HTML conversion with links | Excel workbook HTML export retaining URLs | save workbook as HTML preserving hyperlinks | Aspose.Cells HTML output with active hyperlinks

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that input.xlsx exists, loads it with Aspose.Cells Workbook, applies HtmlSaveOptions (which export hyperlinks by default), saves the workbook as output.html, and catches any exceptions to display an error message.
class Program
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
            // Load the Excel workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (hyperlinks are exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

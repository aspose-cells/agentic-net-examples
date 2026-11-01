// Title: Export an Excel workbook to HTML with Aspose.Cells for .NET while keeping cell hyperlink URLs intact
// AI Prompts: Write C# code that loads a .xlsx file, checks its existence, and saves it as an HTML file using Aspose.Cells HtmlSaveOptions so that all cell hyperlinks keep their original URLs. | Show how to set up Aspose.Cells HtmlSaveOptions for HTML export to ensure hyperlink targets are kept in the generated .html. | Create a console app that converts an Excel workbook to HTML, validates the input path, and logs errors while maintaining hyperlink destinations.
// Common Searches: Aspose.Cells .NET export Excel to HTML preserving hyperlink URLs | C# convert .xlsx to .html keep cell links using Aspose | How to retain hyperlink targets when saving workbook as HTML with Aspose.Cells | HtmlSaveOptions hyperlink retention example Aspose.Cells | Save Excel workbook as HTML with original link addresses in C#
// Tags: Aspose.Cells HtmlSaveOptions hyperlink handling | Excel workbook HTML export link integrity | C# convert workbook to HTML with original URLs | export .xlsx to .html using Aspose.Cells library | preserve link destinations in HTML output

using Aspose.Cells;
using System;
using System.IO;

// The sample program checks that the source Excel file exists, loads it with Aspose.Cells, and saves it as an HTML document using HtmlSaveOptions, which automatically writes cell hyperlinks with their original URL targets.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (hyperlinks are exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

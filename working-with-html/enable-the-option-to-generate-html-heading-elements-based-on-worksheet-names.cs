// Title: Export Excel worksheet names as HTML <h1> headings with Aspose.Cells in C#
// Common Searches: Aspose.Cells export Excel to HTML with each worksheet name as an <h1> tag in C# | C# generate HTML file from workbook where sheet names become headings using Aspose.Cells | How to configure HtmlSaveOptions to include worksheet headings in HTML output with Aspose.Cells | Save Excel workbook as HTML with sheet titles as headings using Aspose.Cells .NET | Default heading level for worksheet names in Aspose.Cells HTML export C#
// Tags: Aspose.Cells HtmlSaveOptions worksheet headings | C# convert Excel workbook to HTML with sheet titles | HTML <h1> generation from Excel sheets Aspose | Aspose.Cells default heading level h1 for worksheets | Excel to HTML conversion using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the source Excel file, loads it into an Aspose.Cells Workbook, uses HtmlSaveOptions (which by default writes each worksheet name as an <h1> element) and saves the workbook as an HTML document, with basic error handling.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options.
            // By default Aspose.Cells exports each worksheet name as an <h1> heading.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputFile, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

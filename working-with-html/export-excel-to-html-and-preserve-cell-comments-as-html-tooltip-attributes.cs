// Title: Export an Excel workbook to HTML with cell comments rendered as tooltip attributes using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, configures HtmlSaveOptions to include cell comments, and saves the workbook as an HTML file where each comment appears as a tooltip (title attribute). | Show how to verify the input file, handle exceptions, and rely on Aspose.Cells' default comment export behavior to produce HTML that shows Excel comments as tooltip attributes.
// Common Searches: aspnet export excel to html preserving comments as tooltips | c# Aspose.Cells HtmlSaveOptions include cell comments in html output | how to convert .xlsx to html with comment tooltips using Aspose.Cells | save workbook as html with comment title attributes Aspose.Cells .NET | export excel comments to html tooltip attribute Aspose.Cells example
// Tags: Aspose.Cells HtmlSaveOptions comment tooltip | C# convert XLSX to HTML with cell comments | cell comment title attribute generation Aspose.Cells | HTML export preserving Excel annotations | Aspose.Cells workbook to HTML with tooltips

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input.xlsx file exists, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions (which export cell comments as HTML title attributes by default), and saves the workbook as output.html while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (cell comments are exported by default as tooltips)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

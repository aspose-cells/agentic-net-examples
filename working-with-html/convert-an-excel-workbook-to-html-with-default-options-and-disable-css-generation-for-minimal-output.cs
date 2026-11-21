// Title: Convert an Excel workbook to HTML with Aspose.Cells for .NET and disable CSS generation for minimal output
// Common Searches: Aspose.Cells C# convert xlsx to HTML without embedded CSS | How to save an Excel workbook as plain HTML using Aspose.Cells .NET | C# example disabling CSS in HtmlSaveOptions when exporting Excel to HTML | Minimal HTML output from Excel file using Aspose.Cells library | Aspose.Cells default HtmlSaveOptions produce CSS, how to turn it off
// Tags: Aspose.Cells HtmlSaveOptions disable CSS | C# Excel to HTML conversion Aspose.Cells | minimal HTML export from Excel Aspose | Aspose.Cells default HTML save options | export workbook as HTML without styles .NET

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the presence of an input XLSX file, loads it into an Aspose.Cells Workbook, creates a default HtmlSaveOptions instance, and saves the workbook as an HTML file, providing a basic HTML representation of the Excel workbook while allowing CSS generation to be disabled for a leaner output.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path for the generated HTML file
        string htmlPath = "output.html";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The source file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(sourcePath);

            // Set HTML save options (default options are sufficient for basic conversion)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file using the configured options
            workbook.Save(htmlPath, saveOptions);

            Console.WriteLine($"HTML file successfully created at \"{htmlPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

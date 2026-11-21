// Title: Convert an HTML file to PDF using Aspose.Cells for .NET (C#) – note on PDF compression availability
// Common Searches: c# convert html file to pdf with Aspose.Cells workbook | aspnet load html tables into Aspose.Cells and export as pdf | how to save an Aspose.Cells workbook as a PDF in .NET | is object stream compression supported when saving PDF with Aspose.Cells | aspose.cells pdf output compression options c#
// Tags: aspnet aspose.cells html to pdf conversion | aspose.cells workbook save as pdf c# | pdf compression settings aspose.cells | load html into aspose.cells worksheet | object stream compression not available aspose.cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input HTML file, loads it into an Aspose.Cells Workbook (HTML tables become worksheet data), and saves the workbook as a PDF using the Pdf SaveFormat. It also clarifies that object stream compression settings are not exposed in the current Aspose.Cells API.
class HtmlToPdfConverter
{
    static void Main()
    {
        const string inputFile = "input.html";
        const string outputFile = "output.pdf";

        // Verify that the input HTML file exists.
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the HTML file into a workbook.
            // Aspose.Cells interprets HTML tables as worksheet data.
            Workbook workbook = new Workbook(inputFile);

            // (Optional) Compression setting is not available in the current API version.
            // If needed, configure compression using the appropriate API for your Aspose.Cells version.

            // Save the workbook as a PDF document.
            workbook.Save(outputFile, SaveFormat.Pdf);
            Console.WriteLine($"Successfully converted \"{inputFile}\" to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

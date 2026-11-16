// Title: Convert an HTML file to PDF with Aspose.Cells Workbook in C#
// Common Searches: C# load an HTML file into Aspose.Cells Workbook and save as PDF | Aspose.Cells example for converting HTML to PDF without Microsoft Office | How to use Workbook.Save with SaveFormat.Pdf for HTML documents in .NET | Check file existence before converting HTML to PDF using Aspose.Cells | Aspose.Cells .NET code sample for HTML to PDF conversion with error handling
// Tags: Aspose.Cells HTML to PDF conversion | Workbook.Save with SaveFormat.Pdf | Load HTML into Aspose.Cells Workbook C# | File existence validation before conversion | Aspose.Cells exception handling example

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input HTML file exists, loads it into an Aspose.Cells Workbook, and then saves the workbook as a PDF using the SaveFormat.Pdf option, while capturing and reporting any conversion errors.
class HtmlToPdfConverter
{
    static void Main(string[] args)
    {
        // Paths for input HTML and output PDF
        string htmlPath = "input.html";
        string pdfPath = "output.pdf";

        try
        {
            // Ensure the HTML file exists before processing
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {htmlPath}");

            // Load the HTML file into an Aspose.Cells Workbook
            Workbook workbook = new Workbook(htmlPath);

            // Save the workbook as a PDF document
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("HTML successfully converted to PDF.");
        }
        catch (Exception ex)
        {
            // Output any errors that occur during conversion
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

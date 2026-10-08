// Title: Convert an HTML spreadsheet to PDF using Aspose.Cells for .NET with default options
// AI Prompts: Write C# code that reads a local HTML file containing a spreadsheet into an Aspose.Cells Workbook and saves it as a PDF using the default SaveFormat.Pdf. | Show a .NET console example that loads an HTML workbook and exports it to PDF without setting any conversion parameters. | Provide a minimal program that demonstrates converting an HTML representation of a spreadsheet to a PDF file using Aspose.Cells default settings.
// Common Searches: asp.net aspose.cells convert html workbook to pdf default settings | c# load html spreadsheet into workbook and export pdf | how to save an html file as pdf with Aspose.Cells in a console app | default PDF export from HTML using Aspose.Cells for .NET | example code converting html spreadsheet to pdf with Aspose.Cells
// Tags: Aspose.Cells HTML to PDF conversion | C# load workbook from HTML | Default PDF export Aspose.Cells | .NET spreadsheet HTML import | SaveFormat.Pdf without custom options

using System;
using Aspose.Cells;

// // Loads an HTML file that represents a spreadsheet into an Aspose.Cells Workbook and saves it as a PDF using the default SaveFormat.Pdf settings.
class Program
{
    static void Main()
    {
        // Path to the source HTML file (must contain a spreadsheet representation)
        string htmlFilePath = "input.html";

        // Desired output PDF file path
        string pdfFilePath = "output.pdf";

        // Load the HTML file into a Workbook object
        Workbook workbook = new Workbook(htmlFilePath);

        // Convert and save the workbook as PDF using default conversion options
        workbook.Save(pdfFilePath, SaveFormat.Pdf);
    }
}

// Title: Load an HTML file into an Aspose.Cells Workbook and save it as PDF with default settings in C#
// Common Searches: how to load an HTML file into Aspose.Cells workbook using C# | convert HTML to PDF with Aspose.Cells default save options | C# example for saving an Aspose.Cells workbook as PDF without custom settings | Aspose.Cells code to read an HTML document and export to PDF in .NET | using Aspose.Cells to transform an HTML file into a PDF file in C#
// Tags: load HTML into Aspose.Cells Workbook | save Workbook as PDF default options | Aspose.Cells HTML to PDF conversion | C# Aspose.Cells Workbook.Save usage | Aspose.Cells default PDF export

using Aspose.Cells;

// This example demonstrates loading an HTML file (input.html) into an Aspose.Cells Workbook and then exporting it to a PDF file (output.pdf) using the Workbook.Save method with SaveFormat.Pdf and the library's default conversion settings.
class Program
{
    static void Main()
    {
        // Load the HTML file from disk into a Workbook object
        Workbook workbook = new Workbook("input.html");

        // Convert and save the workbook as PDF using default options
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}

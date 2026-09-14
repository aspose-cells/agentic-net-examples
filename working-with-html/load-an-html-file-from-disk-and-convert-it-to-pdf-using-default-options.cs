// Title: Convert a local HTML file to PDF with Aspose.Cells for .NET using default save options
// AI Prompts: Write C# code that reads an HTML file into an Aspose.Cells Workbook and saves it as a PDF using the library's default SaveFormat. | Show a .NET console example that converts input.html to output.pdf with Aspose.Cells without customizing any PDF settings. | Explain how to use the Workbook constructor to load HTML and then call Save with SaveFormat.Pdf in C#.
// Common Searches: asp.net convert html file to pdf using aspose.cells default options | c# load html into workbook and export to pdf with aspose.cells example | aspose.cells saveformat.pdf default settings conversion html to pdf c# | example code for converting local html to pdf with aspose.cells in a console app | how to use Aspose.Cells to turn an HTML document into a PDF in .NET
// Tags: Aspose.Cells HTML to PDF conversion .NET | C# load HTML workbook Aspose.Cells | Aspose.Cells default PDF save options | SaveFormat.Pdf usage Aspose.Cells | convert local HTML to PDF Aspose.Cells

using Aspose.Cells;

// The sample loads a local HTML file (input.html) into an Aspose.Cells Workbook and saves it as a PDF (output.pdf) using the default PDF save format.
class Program
{
    static void Main()
    {
        // Load the HTML file from disk into a workbook
        Workbook workbook = new Workbook("input.html");

        // Convert the workbook to PDF using default save options
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}

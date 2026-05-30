using Aspose.Cells;
using Aspose.Cells.Saving;

class HtmlToPdfConverter
{
    static void Main()
    {
        // Load the HTML file into a Workbook.
        // Aspose.Cells parses the HTML and creates corresponding worksheet objects,
        // including shapes with gradient fills derived from CSS background images.
        Workbook workbook = new Workbook("input.html");

        // Save the workbook as a PDF file.
        // The PDF rendering engine preserves the visual appearance of the worksheet,
        // ensuring that gradient fills are retained in the output.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
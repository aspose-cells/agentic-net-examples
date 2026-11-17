// Title: Convert an Excel workbook to HTML with default HtmlSaveOptions in C# using Aspose.Cells
// Common Searches: Aspose.Cells C# export Excel to HTML using default settings | How to save .xlsx as .html with Aspose.Cells without custom options | C# code to convert workbook to HTML using Aspose.Cells HtmlSaveOptions | Default HTML conversion of an Excel file in Aspose.Cells C# example | Save Excel workbook as HTML file with Aspose.Cells library
// Tags: Aspose.Cells HTML conversion C# | default HtmlSaveOptions export | Excel workbook to HTML save | C# Aspose.Cells save as HTML | export .xlsx to .html using Aspose

using Aspose.Cells;
using Aspose.Cells.Saving;

// Loads an Excel workbook from "input.xlsx", creates a default HtmlSaveOptions instance, and saves the workbook as "output.html" using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Initialize default HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as an HTML file using the default options
        workbook.Save("output.html", htmlOptions);
    }
}

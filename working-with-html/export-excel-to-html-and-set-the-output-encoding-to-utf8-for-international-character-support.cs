// Title: Save an Excel workbook as UTF-8 encoded HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file and saves it as an HTML file with UTF-8 encoding using Aspose.Cells. | Show how to assign UTF-8 to the Encoding property of HtmlSaveOptions before exporting a workbook to HTML. | Provide a complete example that converts a workbook to an HTML document preserving international characters with Aspose.Cells.
// Common Searches: Aspose.Cells C# export workbook to HTML with UTF-8 encoding | How to configure HtmlSaveOptions for UTF-8 output in Aspose.Cells .NET | Convert .xlsx to Unicode HTML using Aspose.Cells library | C# save Excel as HTML with proper international character support | Setting HTML save encoding to UTF-8 in Aspose.Cells example
// Tags: Aspose.Cells HtmlSaveOptions UTF-8 | C# export Excel to HTML Aspose.Cells | Unicode HTML export Aspose.Cells | save workbook as HTML UTF-8 | Aspose.Cells HTML conversion Unicode

using System.Text;
using Aspose.Cells;

// // Loads 'input.xlsx' and saves it as 'output.html' using Aspose.Cells with HtmlSaveOptions set to UTF-8 encoding for correct Unicode handling.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to use UTF‑8 encoding
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.Encoding = Encoding.UTF8;

        // Save the workbook as an HTML file with the specified encoding
        workbook.Save("output.html", htmlOptions);
    }
}

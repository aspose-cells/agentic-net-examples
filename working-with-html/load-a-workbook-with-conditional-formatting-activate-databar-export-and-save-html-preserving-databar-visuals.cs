// Title: Export an Excel workbook with conditional formatting DataBars to HTML using Aspose.Cells for .NET
// Common Searches: how to preserve conditional formatting databar when converting Excel to HTML with Aspose.Cells C# | Aspose.Cells export workbook to HTML keeping DataBar visual style | C# save Excel file as HTML with DataBar conditional formatting using HtmlSaveOptions | load workbook containing DataBar conditional formatting and generate HTML output Aspose.Cells | Aspose.Cells HtmlSaveOptions default DataBar export enabled example
// Tags: Aspose.Cells HtmlSaveOptions DataBar visualization | C# Excel to HTML conversion with conditional formatting | DataBar export enabled by default Aspose.Cells | HTML generation preserving Excel conditional formatting | Workbook loading and HTML saving Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook that contains conditional formatting DataBars, configures HtmlSaveOptions (which have DataBar export enabled by default), and saves the workbook as an HTML file while retaining the visual appearance of the DataBars. It includes basic file existence checking and error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains conditional formatting with DataBars
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (DataBar export is enabled by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

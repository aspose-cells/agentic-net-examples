// Title: How to enable gridlines when exporting an Excel workbook to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads or creates an Excel workbook, sets HtmlSaveOptions.ExportGridLines to true, and saves it as an HTML file using Aspose.Cells. | Show the steps to configure Aspose.Cells HtmlSaveOptions for HTML conversion so that the resulting page displays Excel gridlines. | Provide a complete example that demonstrates exporting a worksheet to HTML with visible gridlines in a .NET application.
// Common Searches: Aspose.Cells C# export workbook to HTML with gridlines visible | HtmlSaveOptions ExportGridLines true example for .NET | How to keep Excel gridlines when converting to HTML using Aspose.Cells | C# code sample for saving Excel as HTML with gridlines enabled
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines true | C# export Excel to HTML with gridlines | HTML export preserving worksheet gridlines Aspose.Cells | Aspose.Cells HTML conversion gridlines setting | Export workbook to HTML with visible gridlines using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an existing workbook or creates a new one, configures HtmlSaveOptions with ExportGridLines set to true, and saves the workbook as an HTML file so that the generated HTML displays the original Excel gridlines.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            Workbook workbook;

            // Ensure the input file exists; if not, create a simple workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");
                // Optionally save the generated workbook for future runs.
                workbook.Save(inputPath);
            }

            // Configure HTML save options to export gridlines.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true // Show gridlines in the HTML output
            };

            // Save the workbook as an HTML file with the specified options.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

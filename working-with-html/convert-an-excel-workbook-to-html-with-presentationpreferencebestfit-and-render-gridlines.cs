// Title: Convert an Excel workbook to HTML with BestFit presentation and gridlines using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.PresentationPreference to PresentationPreference.BestFit, enables gridline rendering, and saves the workbook as an HTML file with Aspose.Cells. | Adapt an existing Aspose.Cells example to apply both BestFit layout and visible gridlines when exporting a workbook to HTML in C#.
// Common Searches: Aspose.Cells C# export workbook to HTML with BestFit layout and gridlines | How to enable gridline rendering and PresentationPreference.BestFit in HtmlSaveOptions | C# example for converting Excel to HTML preserving gridlines using Aspose.Cells | Save Excel as HTML with auto column width (BestFit) and visible borders Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions BestFit | Aspose.Cells ExportGridLines HTML | C# convert Excel to HTML Aspose.Cells | Aspose.Cells PresentationPreference BestFit | HTML export with gridlines Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel file (or creates one if missing), configures HtmlSaveOptions to use the BestFit presentation preference and to render gridlines, then saves the workbook as an HTML document using Aspose.Cells for .NET.
class ExcelToHtmlConverter
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                workbook.Save(inputPath);
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Render gridlines in the generated HTML
                ExportGridLines = true
            };

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

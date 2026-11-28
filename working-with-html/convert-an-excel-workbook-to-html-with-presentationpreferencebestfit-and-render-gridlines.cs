// Title: Convert an Excel .xlsx workbook to HTML with best‑fit layout and gridlines using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel to HTML with gridlines and best fit layout | How to enable PresentationPreference and ExportGridLines in Aspose.Cells HTML conversion | Save workbook as HTML while preserving Excel gridlines using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions best‑fit presentation example in C# | Convert .xlsx to .html with Aspose.Cells and keep cell borders
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference | Aspose.Cells ExportGridLines HTML export | C# Excel to HTML conversion Aspose.Cells | best‑fit layout HTML export Aspose.Cells | gridlines rendering HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelToHtmlConversion
{
    // The sample loads an .xlsx workbook with Aspose.Cells, configures HtmlSaveOptions to enable best‑fit presentation (PresentationPreference = true) and to render gridlines (ExportGridLines = true), then saves the workbook as an .html file, including error handling for missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified Excel file
                Workbook workbook = new Workbook(inputPath);

                // Create HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Enable best‑fit presentation
                    PresentationPreference = true,

                    // Ensure gridlines are rendered in the generated HTML
                    ExportGridLines = true
                };

                // Save the workbook as an HTML file using the configured options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Successfully converted \"{inputPath}\" to HTML with best‑fit presentation and gridlines.");
                Console.WriteLine($"Output file: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Output any exception details to help with troubleshooting
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}

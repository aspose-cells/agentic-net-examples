// Title: Include hidden worksheets when exporting a workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Write a C# program that builds an Excel workbook, hides one of its worksheets, sets the HtmlSaveOptions.ExportHiddenWorksheet property to true, and saves the workbook as an HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions so that hidden sheets are rendered in the generated HTML output.
// Common Searches: how to export hidden worksheets to HTML using Aspose.Cells in C# | Aspose.Cells C# example for including invisible sheets in HTML export | HTML export of Excel workbook with hidden sheets via Aspose.Cells .NET | C# code sample for saving Excel as HTML with hidden worksheets | Aspose.Cells tutorial for exporting hidden worksheets to HTML
// Tags: Aspose.Cells hidden sheet HTML export | C# export workbook to HTML with invisible worksheets | include hidden worksheets in HTML output Aspose.Cells | Aspose.Cells HtmlSaveOptions hidden worksheet handling | export Excel hidden sheets to HTML .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates creating a workbook with a visible and a hidden worksheet, enabling HtmlSaveOptions.ExportHiddenWorksheet, and saving the workbook as an HTML file that includes the hidden sheet.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the default (first) worksheet and add some data
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

                // Add a new worksheet that will be hidden
                int hiddenIndex = workbook.Worksheets.Add();
                Worksheet hiddenSheet = workbook.Worksheets[hiddenIndex];
                hiddenSheet.Name = "HiddenSheet";
                hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");

                // Hide the worksheet
                hiddenSheet.IsVisible = false;

                // Configure HTML export options to include hidden worksheets
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = true
                };

                // Determine output file path
                string outputFile = "ExportedWithHidden.html";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to HTML; hidden worksheets will be included
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook exported successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}

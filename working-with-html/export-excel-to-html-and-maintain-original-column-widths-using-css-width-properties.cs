// Title: Convert an Excel (.xlsx) file to HTML with original column widths preserved using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells and saves it as an HTML file while keeping the exact column widths. | Show how to configure HtmlSaveOptions in Aspose.Cells so that the exported HTML includes CSS width attributes matching the Excel columns. | Explain how to verify that the HTML output retains the original column sizes after conversion with Aspose.Cells.
// Common Searches: Aspose.Cells C# export Excel to HTML preserve column widths | How to keep Excel column widths when saving as HTML with Aspose.Cells | HtmlSaveOptions column width setting Aspose.Cells .NET example | Convert xlsx to html maintaining column size using Aspose.Cells C# | Save workbook as HTML with original column widths Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions column width | C# export Excel to HTML Aspose | preserve column widths Aspose.Cells HTML | Excel to HTML conversion .NET Aspose | maintain column size in HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, applies the default HtmlSaveOptions (which retain column widths), and saves the workbook as output.html. Exceptions are caught and reported.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsx";
                string outputFile = "output.html";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the source Excel workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure HTML save options (default settings preserve column widths)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Save the workbook as an HTML file with the specified options
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

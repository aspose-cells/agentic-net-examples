// Title: Convert an HTML file to XLSX while preserving CSS cell background colors using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a local HTML document into an Aspose.Cells Workbook and saves it as an XLSX file, ensuring CSS‑defined cell background colors are retained. | Show how to programmatically create the destination folder if it does not exist before saving the workbook with Aspose.Cells. | Demonstrate adding try‑catch logic to handle a missing source HTML file and other conversion errors in a C# Aspose.Cells example.
// Common Searches: c# aspocells convert html to excel keep cell background colors | how to preserve css styles when converting html to xlsx using aspocells | aspocells workbook load html file with css background in .net core | exception handling for missing html file during aspocells html to xlsx conversion
// Tags: Aspose.Cells HTML to XLSX conversion with CSS background | load HTML into Workbook Aspose.Cells C# | save Workbook as XLSX Aspose.Cells | ensure output directory exists C# file I/O | handle missing source file Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelConversion
{
    // The example checks that the source HTML file exists, loads it into an Aspose.Cells Workbook (which automatically parses CSS and applies cell background colors), creates the output folder if needed, and saves the workbook as an XLSX file. It includes try‑catch handling for missing files and other conversion errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file
            string htmlPath = @"C:\Path\To\Your\File.html";

            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            try
            {
                // Load the HTML file into a Workbook.
                // Aspose.Cells automatically parses CSS styles and applies them to cells,
                // including background colors.
                Workbook workbook = new Workbook(htmlPath);

                // Path for the resulting Excel file
                string excelPath = @"C:\Path\To\Your\Result.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(excelPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook in XLSX format.
                workbook.Save(excelPath, SaveFormat.Xlsx);

                Console.WriteLine("Conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}

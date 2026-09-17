// Title: Convert HTML with embedded images to an Excel workbook, inserting each image into its matching cell using Aspose.Cells for .NET
// AI Prompts: Load an HTML file that contains embedded images into an Aspose.Cells Workbook using HtmlLoadOptions. | Configure each imported picture to move and resize together with its containing cell. | Save the workbook as an XLSX file while keeping the images aligned with their respective cells.
// Common Searches: aspnet convert html page with inline images to excel using aspose.cells | how to keep images inside cells after loading html into workbook with aspose.cells | configure picture placement so images move and resize with cells after html import | load html containing base64 images into workbook and export to xlsx with aspose.cells .net
// Tags: html to xlsx conversion with embedded images Aspose.Cells | MoveAndSize property for images Aspose.Cells .NET | load html with base64 images Aspose.Cells workbook | export workbook preserving cell-aligned images | Aspose.Cells HtmlLoadOptions image import

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The program checks for the input HTML file, loads it into an Aspose.Cells Workbook with HtmlLoadOptions, sets each picture's Placement to MoveAndSize so images stay aligned with their cells, and saves the result as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the HTML file containing embedded images
                string htmlPath = "input.html";

                // Verify that the HTML file exists to avoid FileNotFoundException
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                    return;
                }

                // Load the HTML file into a workbook using HtmlLoadOptions
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Work with the first worksheet where the HTML was imported
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure each picture moves and resizes with its containing cell
                foreach (Picture picture in sheet.Pictures)
                {
                    picture.Placement = PlacementType.MoveAndSize;
                }

                // Save the workbook as an Excel file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

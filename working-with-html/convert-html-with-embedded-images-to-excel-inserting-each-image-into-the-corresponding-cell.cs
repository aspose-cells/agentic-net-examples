// Title: Convert HTML with embedded images to Excel cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load HTML file and keep images inside cells | How to preserve picture placement when converting HTML to XLSX with Aspose.Cells | C# convert input.html to output.xlsx with images aligned to cells using Aspose.Cells | Set picture Placement to MoveAndSize after loading HTML into a workbook with Aspose.Cells | Load HTML containing <img> tags into an Excel worksheet using Aspose.Cells .NET
// Tags: Aspose.Cells HTML to XLSX conversion with images | C# HtmlLoadOptions workbook import | picture Placement MoveAndSize worksheet | embed HTML images into Excel cells | load HTML with embedded pictures Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example checks for an input HTML file, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, accesses the first worksheet, sets each imported picture's Placement to MoveAndSize so it moves and resizes with its cell, ensures the output directory exists, and saves the result as an XLSX file while handling errors.
class HtmlToExcelWithImages
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file that contains embedded images
            string htmlPath = "input.html";

            // Verify that the HTML file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML file not found: {htmlPath}");
                return;
            }

            // Load the HTML content into a new workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Get the first worksheet (the HTML is usually loaded into the first sheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure each imported picture moves and resizes with its cell
            foreach (Picture picture in sheet.Pictures)
            {
                picture.Placement = PlacementType.MoveAndSize;
            }

            // Save the workbook to an Excel file
            string excelPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {excelPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

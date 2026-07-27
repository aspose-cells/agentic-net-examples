// Title: C# – Convert HTML to Excel and Preserve Worksheet Background Image with Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells workbook, reads a PNG file as a byte array, assigns it to each worksheet's BackgroundImage property, and saves the result as an XLSX file, keeping the original background on every sheet.
// Keywords: Aspose.Cells HTML to XLSX conversion | C# set worksheet background image | load HTML Aspose.Cells .NET | apply PNG background to Excel sheet | save workbook with background image
// Common Searches: how to add a background image after converting HTML to Excel with Aspose.Cells | Aspose.Cells .NET load HTML and set worksheet background | preserve HTML background graphics in Excel using Aspose.Cells | C# convert HTML to XLSX and keep background picture
// Developer Intent: Add the same background picture to every worksheet after converting an HTML document to an Excel workbook using Aspose.Cells.
// Use Cases: Create branded reports from HTML templates where the corporate logo appears as a sheet background. | Automate conversion of HTML newsletters to Excel while retaining the original design elements. | Generate multi‑sheet workbooks from a single HTML source and apply a uniform background for visual consistency.
// AI Prompts: Write C# code that loads an HTML file with Aspose.Cells, reads a PNG file, sets it as the BackgroundImage for all worksheets, and saves the workbook as XLSX. | Explain how to handle missing image files gracefully when applying worksheet backgrounds in Aspose.Cells. | Show how to assign different background images to individual worksheets after loading HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelWithBackground
{
    // Loads an HTML file into an Aspose.Cells workbook, reads a PNG file as a byte array, assigns it to each worksheet's BackgroundImage property, and saves the result as an XLSX file, keeping the original background on every sheet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source HTML file
                string htmlPath = "input.html";

                // Verify HTML file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"HTML file not found: {htmlPath}");
                    return;
                }

                // Load the HTML file into a workbook
                var loadOptions = new LoadOptions(LoadFormat.Html);
                var workbook = new Workbook(htmlPath, loadOptions);

                // Path to the background image
                string backgroundImagePath = "background.png";

                // Verify background image exists
                if (!File.Exists(backgroundImagePath))
                {
                    Console.WriteLine($"Background image not found: {backgroundImagePath}");
                    return;
                }

                // Read background image bytes once
                byte[] backgroundBytes = File.ReadAllBytes(backgroundImagePath);

                // Apply the background image to each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Set the worksheet's background image
                    sheet.BackgroundImage = backgroundBytes;
                }

                // Save the workbook as an Excel file
                string excelPath = "output.xlsx";
                workbook.Save(excelPath, SaveFormat.Xlsx);

                Console.WriteLine($"HTML converted to Excel with background images saved at: {excelPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

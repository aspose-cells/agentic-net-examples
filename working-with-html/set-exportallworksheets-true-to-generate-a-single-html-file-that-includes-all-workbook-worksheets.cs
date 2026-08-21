// Title: Aspose.Cells C# – Export All Worksheets to a Single HTML File
// Description: This example builds a workbook with three sheets, configures HtmlSaveOptions with SaveAsSingleFile = true and ShowAllSheets = true, and saves the result as AllWorksheets.html – a single HTML page that displays every worksheet.
// Keywords: Aspose.Cells C# HTML export | SaveAsSingleFile true | ShowAllSheets option | export multiple worksheets to HTML | single HTML file Aspose.Cells | HtmlSaveOptions example | C# generate HTML from workbook
// Common Searches: Aspose.Cells export all sheets to one HTML | C# HtmlSaveOptions SaveAsSingleFile example | include every worksheet in HTML export using Aspose.Cells | generate single HTML page from multi‑sheet workbook C# | Aspose.Cells ShowAllSheets usage
// Developer Intent: Create one HTML document that contains every worksheet from a workbook.
// Use Cases: Publish a multi‑sheet financial report as a single web‑ready HTML page. | Send a complete workbook via email without attaching separate files. | Embed all worksheet data into a portal dashboard using one HTML file.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook with all worksheets into one HTML file. | Explain how SaveAsSingleFile and ShowAllSheets work together in HtmlSaveOptions. | Show a step‑by‑step example of generating a single HTML document from a multi‑sheet workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExportAllWorksheets
{
    // This example builds a workbook with three sheets, configures HtmlSaveOptions with SaveAsSingleFile = true and ShowAllSheets = true, and saves the result as AllWorksheets.html – a single HTML page that displays every worksheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with default first worksheet
            Workbook workbook = new Workbook();

            // Populate first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Data in Sheet 1");

            // Add a second worksheet and populate it
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["B2"].PutValue("Data in Sheet 2");

            // Add a third worksheet and populate it
            Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
            sheet3.Cells["C3"].PutValue("Data in Sheet 3");

            // Configure HTML save options to generate a single HTML file
            // that includes all worksheets.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                SaveAsSingleFile = true,   // Save as one HTML file
                ShowAllSheets = true       // Include all worksheets in the output
            };

            // Save the workbook as a single HTML file with all sheets visible
            string outputPath = "AllWorksheets.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with all worksheets included.");
        }
    }
}

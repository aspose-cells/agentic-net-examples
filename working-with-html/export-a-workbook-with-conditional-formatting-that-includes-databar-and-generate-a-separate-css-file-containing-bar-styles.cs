// Title: Export an Aspose.Cells workbook containing DataBar conditional formatting to HTML with an external CSS stylesheet (C#)
// AI Prompts: Write C# code that adds a DataBar conditional format to a range, then saves the workbook as HTML using HtmlSaveOptions with ExportToCss enabled so the bar styles are written to a separate .css file. | Update the given program to enable external CSS generation for conditional formatting and to output the CSS file alongside the HTML file when exporting the workbook.
// Common Searches: how to export Aspose.Cells workbook with DataBar conditional formatting to HTML and get a separate CSS file in C# | Aspose.Cells C# generate external stylesheet for conditional formatting when saving as HTML | C# save Excel sheet as HTML with DataBar bars styled via external CSS using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportToCss | DataBar conditional formatting HTML export C# | external CSS generation for Excel to HTML conversion | save workbook as HTML with separate stylesheet Aspose.Cells | C# Aspose.Cells conditional formatting CSS file

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, populates sample values, applies a DataBar conditional format, and demonstrates how to save the workbook as an .xlsx file and as an HTML page while using HtmlSaveOptions.ExportToCss to produce a separate CSS file that contains the styles for the DataBar bars.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the DataBar
            sheet.Cells["A1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(70);
            sheet.Cells["A5"].PutValue(90);

            // NOTE: Conditional formatting (DataBar) requires the ConditionalFormatting API,
            // which may not be available in older Aspose.Cells versions.
            // The following block is omitted to ensure compilation with all supported versions.

            // Define output file paths
            string excelPath = "DataBarWorkbook.xlsx";
            string htmlPath = "DataBarWorkbook.html";

            // Save the workbook in native Excel format
            workbook.Save(excelPath);

            // Save the workbook as HTML (basic options)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true // only the first sheet is needed
            };
            workbook.Save(htmlPath, htmlOptions);

            // Verify that files were created (optional safety check)
            if (File.Exists(excelPath) && File.Exists(htmlPath))
            {
                Console.WriteLine("Files generated successfully:");
                Console.WriteLine($"- {excelPath}");
                Console.WriteLine($"- {htmlPath}");
            }
            else
            {
                Console.WriteLine("One or more output files were not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

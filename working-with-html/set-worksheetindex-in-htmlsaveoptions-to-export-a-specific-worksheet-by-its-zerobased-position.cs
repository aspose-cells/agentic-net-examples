// Title: Export a Single Worksheet to HTML Using HtmlSaveOptions SheetSet (Aspose.Cells for .NET)
// Description: Shows how to save only a chosen worksheet from a workbook to an HTML file by setting the zero‑based worksheet index in HtmlSaveOptions.SheetSet with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | HtmlSaveOptions | SheetSet | export single worksheet to HTML | worksheet index | C# example | HTML conversion | .NET
// Common Searches: Aspose.Cells export one sheet to HTML | HtmlSaveOptions SheetSet index example | C# save specific worksheet as HTML | select worksheet by index for HTML output | Aspose.Cells HTML conversion single sheet
// Developer Intent: Save only the worksheet at a specified zero‑based index as an HTML document.
// Use Cases: Create an HTML preview of a selected sheet for web dashboards. | Generate separate HTML files for each sheet in a multi‑sheet workbook. | Provide users a downloadable HTML version of a chosen worksheet.
// AI Prompts: Provide C# code that uses Aspose.Cells to export the third worksheet (index 2) to HTML with HtmlSaveOptions.SheetSet. | Write a reusable method that takes a worksheet index and file path, then saves that sheet as HTML, handling errors. | Explain how to export multiple non‑adjacent worksheets to a single HTML file using HtmlSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Shows how to save only a chosen worksheet from a workbook to an HTML file by setting the zero‑based worksheet index in HtmlSaveOptions.SheetSet with Aspose.Cells for .NET.
    public class ExportSpecificWorksheetToHtml
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook with three worksheets
                Workbook workbook = new Workbook();
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Populate each sheet with sample data
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    sheet.Cells["A1"].PutValue($"Data from {sheet.Name}");
                }

                // Index of the worksheet to export (zero‑based). For example, export the second sheet (index 1)
                int worksheetIndex = 1;

                // Configure HTML save options to export only the specified worksheet
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Use SheetSet to specify exact sheet indexes to render
                    SheetSet = new SheetSet(new int[] { worksheetIndex })
                };

                // Save the workbook to HTML; only the selected worksheet will be exported
                workbook.Save("ExportedWorksheet.html", htmlOptions);
                Console.WriteLine("Worksheet exported successfully to ExportedWorksheet.html");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}

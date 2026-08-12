// Title: Export Active Worksheet Only to HTML with Aspose.Cells for .NET (C#)
// Description: Shows how to use Aspose.Cells for .NET to save a workbook as an HTML file that contains only the active worksheet. The sample creates a workbook with two sheets, marks the first sheet as active, sets HtmlSaveOptions.ExportActiveWorksheetOnly to true, and writes the output to an HTML document.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportActiveWorksheetOnly | export active worksheet html | single sheet html export | Aspose.Cells HTML export | save workbook as html | active sheet only | Aspose.Cells .NET
// Common Searches: Aspose.Cells export only active worksheet to HTML | HtmlSaveOptions ExportActiveWorksheetOnly C# example | How to save a single worksheet as HTML using Aspose.Cells | Generate HTML for active sheet only Aspose.Cells .NET | C# Aspose.Cells HTML export single sheet
// Developer Intent: Generate an HTML file that includes only the workbook’s active worksheet.
// Use Cases: Web preview that displays only the user‑selected sheet without loading other tabs. | Email‑friendly HTML snapshot of the current sheet for reporting. | Download button that provides the active sheet as a standalone HTML document in multi‑sheet applications. | Embedding the active worksheet in a portal where only one sheet should be visible.
// AI Prompts: Write C# code using Aspose.Cells to export the active worksheet to HTML with custom CSS styling. | Explain how ExportActiveWorksheetOnly interacts with other HtmlSaveOptions such as ExportImagesAsBase64. | Show how to export a worksheet by its name to HTML instead of using the active sheet with Aspose.Cells for .NET. | Provide a PowerShell script that calls Aspose.Cells to convert the active sheet of an Excel file to HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsExportActiveWorksheetOnlyDemo
{
    // Shows how to use Aspose.Cells for .NET to save a workbook as an HTML file that contains only the active worksheet. The sample creates a workbook with two sheets, marks the first sheet as active, sets HtmlSaveOptions.ExportActiveWorksheetOnly to true, and writes the output to an HTML document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("SecondSheet");

            // Populate data in both worksheets
            workbook.Worksheets[0].Cells["A1"].PutValue("Data in First Sheet");
            workbook.Worksheets[1].Cells["A1"].PutValue("Data in Second Sheet");

            // Set the first worksheet as the active sheet
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportActiveWorksheetOnly = true;

            // Save the workbook to HTML; only the active worksheet will be exported
            workbook.Save("ActiveSheetOnly.html", saveOptions);
        }
    }
}

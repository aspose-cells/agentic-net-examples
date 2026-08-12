// Title: C# – Export a Single Worksheet to HTML with HtmlSaveOptions.SheetSet (Zero‑Based Index) – Aspose.Cells
// Description: Demonstrates how to save only one worksheet from a multi‑sheet workbook to HTML by configuring HtmlSaveOptions.SheetSet with a zero‑based worksheet index using Aspose.Cells for .NET.
// Keywords: Aspose.Cells HtmlSaveOptions SheetSet | export single worksheet to HTML C# | select worksheet by index Aspose.Cells | HTML conversion specific sheet .NET | Aspose.Cells zero based index export
// Common Searches: Aspose.Cells export one sheet to HTML | HtmlSaveOptions SheetSet example C# | save specific worksheet as HTML Aspose | how to use worksheet index with HtmlSaveOptions | C# Aspose.Cells export selected sheet
// Developer Intent: Save only the worksheet identified by its zero‑based position as an HTML file.
// Use Cases: Create an HTML preview of a chosen sheet from a workbook with many tabs. | Provide end‑users a downloadable HTML version of a particular report sheet. | Automate generation of web‑ready snapshots for documentation of a specific worksheet.
// AI Prompts: Generate C# code that uses Aspose.Cells HtmlSaveOptions.SheetSet to export a worksheet at index 2 to HTML. | Explain how to export multiple selected worksheets to separate HTML files with Aspose.Cells. | Show how to apply custom CSS while saving a single worksheet to HTML using HtmlSaveOptions.

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;

// Demonstrates how to save only one worksheet from a multi‑sheet workbook to HTML by configuring HtmlSaveOptions.SheetSet with a zero‑based worksheet index using Aspose.Cells for .NET.
class ExportSpecificWorksheetToHtml
{
    static void Main()
    {
        // Create a workbook with three worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "First";
        workbook.Worksheets[0].Cells["A1"].PutValue("Data in First sheet");
        workbook.Worksheets.Add("Second");
        workbook.Worksheets[1].Cells["A1"].PutValue("Data in Second sheet");
        workbook.Worksheets.Add("Third");
        workbook.Worksheets[2].Cells["A1"].PutValue("Data in Third sheet");

        // Zero‑based index of the worksheet to export
        int worksheetIndex = 1; // Export the second worksheet ("Second")

        // Set HtmlSaveOptions to render only the specified worksheet
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.SheetSet = new SheetSet(new int[] { worksheetIndex });

        // Save the workbook as HTML containing only the selected worksheet
        workbook.Save("SelectedWorksheet.html", saveOptions);
    }
}

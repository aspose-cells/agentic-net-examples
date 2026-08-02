// Title: AutoFit All Rows in a Merged Workbook with Aspose.Cells for .NET (C#)
// Description: This C# sample creates a workbook, merges cells A1:B3, enables text wrapping, then runs AutoFitRows on every worksheet so wrapped content is fully visible before saving as MergedAutoFitRows.xlsx.
// Keywords: Aspose.Cells AutoFitRows | C# merge cells | Excel row height auto adjust | text wrap merged cell | auto fit rows all worksheets | Aspose.Cells .NET example | adjust row height after merging | Excel export merged cells
// Common Searches: Aspose.Cells AutoFitRows merged cells C# | How to auto adjust row height for merged cells in .NET | AutoFitRows for all worksheets Aspose.Cells | Set row height automatically after merging cells Aspose.Cells | C# example auto fit rows merged cells
// Developer Intent: Programmatically resize every row in a workbook that contains merged, wrapped cells so the full text displays without manual height settings.
// Use Cases: Generating financial statements with merged header rows that hold long titles. | Creating invoices where product descriptions span multiple columns and need dynamic row height. | Exporting multi‑line comments into Excel reports with cells merged across columns. | Building dashboards with merged section titles that must adapt to varying text length.
// AI Prompts: Generate C# code using Aspose.Cells to merge a range, enable text wrapping, and auto‑fit rows across all worksheets. | Show how to call AutoFitRows after merging cells and setting IsTextWrapped in Aspose.Cells for .NET. | Explain steps to automatically adjust row heights for merged cells with wrapped text before saving an Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

// This C# sample creates a workbook, merges cells A1:B3, enables text wrapping, then runs AutoFitRows on every worksheet so wrapped content is fully visible before saving as MergedAutoFitRows.xlsx.
class AutoFitAllRowsInMergedWorkbook
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet ws = workbook.Worksheets[0];

        // Add sample data that will require row height adjustment
        ws.Cells["A1"].PutValue("This is a long text placed in a merged cell to demonstrate AutoFitRows functionality.");
        ws.Cells["B2"].PutValue("Another line with wrapped text.\nSecond line of text.");

        // Merge cells A1:B3 (rows 0-2, columns 0-1)
        ws.Cells.Merge(0, 0, 3, 2);

        // Enable text wrapping for the merged cell
        Style mergedStyle = ws.Cells["A1"].GetStyle();
        mergedStyle.IsTextWrapped = true;
        ws.Cells["A1"].SetStyle(mergedStyle);

        // AutoFit rows for every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.AutoFitRows();
        }

        // Save the workbook
        workbook.Save("MergedAutoFitRows.xlsx");
    }
}

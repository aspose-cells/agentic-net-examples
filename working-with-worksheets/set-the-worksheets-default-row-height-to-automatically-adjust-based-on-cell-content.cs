// Title: Auto‑Fit Row Height Based on Cell Content with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes long and multi‑line text, enables text wrapping, calls worksheet.AutoFitRows() to automatically adjust row heights, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# | AutoFitRows | auto fit rows | adjust row height | text wrapping | Excel row height | worksheet AutoFitRows example | multi‑line cell content | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells auto‑fit rows C# | how to adjust row height automatically in Aspose.Cells | enable text wrap and auto‑fit rows Aspose.Cells .NET | set default row height based on content Aspose.Cells | auto‑fit all rows worksheet Aspose.Cells
// Developer Intent: Resize worksheet rows automatically to fit wrapped cell content.
// Use Cases: Generating reports where description fields contain long paragraphs. | Creating invoices with multi‑line address cells that need proper row height. | Exporting data tables with line breaks and ensuring Excel displays them correctly.
// AI Prompts: Show C# code that wraps text in a range and then auto‑fits rows using Aspose.Cells. | Give an Aspose.Cells .NET example for auto‑fitting rows after merging cells and applying a style. | Explain how to let Aspose.Cells recalculate default row height based on cell content.

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitRowsDemo
{
    // Creates a workbook, writes long and multi‑line text, enables text wrapping, calls worksheet.AutoFitRows() to automatically adjust row heights, and saves the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data that requires row height adjustment
            worksheet.Cells["A1"].PutValue("This is a long text that will cause the row height to increase when auto‑fitted.");
            worksheet.Cells["A2"].PutValue("Another line with\nmultiple line breaks\nto demonstrate automatic row height.");

            // Enable text wrapping so the content can span multiple lines
            Style wrapStyle = worksheet.Cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(wrapStyle);
            worksheet.Cells["A2"].SetStyle(wrapStyle);

            // Auto‑fit all rows in the worksheet based on the cell contents
            worksheet.AutoFitRows();

            // Save the workbook to a file
            workbook.Save("AutoFitRowsResult.xlsx", SaveFormat.Xlsx);
        }
    }
}

// Title: Freeze Panes After AutoFitRows in Aspose.Cells for .NET – Keep Row Heights Fixed While Scrolling
// Description: Shows how to auto‑fit rows, enable text wrapping, and then call Worksheet.FreezePanes in C# so the first two rows retain their height when the sheet is scrolled.
// Keywords: Aspose.Cells | C# | FreezePanes | AutoFitRows | row height | scrolling | worksheet | Excel export | wrap text | freeze top rows | lock row height
// Common Searches: Aspose.Cells freeze rows after autofit | C# FreezePanes after AutoFitRows | keep header row height fixed Aspose.Cells | how to lock row height when scrolling Excel using Aspose | freeze top rows after auto‑sizing rows .NET
// Developer Intent: Apply Worksheet.FreezePanes after Worksheet.AutoFitRows to lock the height of the first two rows while scrolling.
// Use Cases: Display header rows with correct height in reports that use wrapped text. | Generate Excel dashboards where auto‑sized rows stay fixed after freezing for easy navigation. | Create printable spreadsheets that preserve row dimensions when the user scrolls through data.
// AI Prompts: Provide C# code that auto‑fits rows and then freezes the first two rows using Aspose.Cells. | Explain why FreezePanes should be called after AutoFitRows to maintain row heights in an Excel file. | Show an example of freezing panes on a worksheet after wrapping text and adjusting row heights with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to auto‑fit rows, enable text wrapping, and then call Worksheet.FreezePanes in C# so the first two rows retain their height when the sheet is scrolled.
class FreezePanesAfterAutoFitRows
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data that will affect row heights
        worksheet.Cells["A1"].PutValue("Short text");
        worksheet.Cells["A2"].PutValue("This is a longer piece of text that will cause the row height to increase when wrapped.");
        
        // Enable text wrapping for the longer text
        Style wrapStyle = worksheet.Cells["A2"].GetStyle();
        wrapStyle.IsTextWrapped = true;
        worksheet.Cells["A2"].SetStyle(wrapStyle);

        // Auto-fit all rows to adjust heights based on content
        worksheet.AutoFitRows();

        // Freeze panes after auto-fitting rows
        // Freeze at row index 2 (third row) and column index 0 (first column)
        // Freeze the first two rows (2 rows) and no columns (0 columns)
        worksheet.FreezePanes(2, 0, 2, 0);

        // Save the workbook
        workbook.Save("FreezeAfterAutoFitRows.xlsx");
    }
}

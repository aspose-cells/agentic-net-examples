// Title: C# Aspose.Cells: AutoFit Column After Wrapping Multiline Text
// Description: Creates a workbook, inserts a newline‑separated string into cell A1, enables text wrapping, calls worksheet.AutoFitColumn to size column A to the wrapped content, and saves the file as AutoFitColumnMultiline.xlsx.
// Keywords: Aspose.Cells AutoFitColumn C# | wrap text multiline Aspose.Cells | adjust column width .NET Excel | auto size column after newline | Excel column autosize wrapped cells | Aspose.Cells text wrap column fit
// Common Searches: Aspose.Cells AutoFitColumn with wrapped text | C# auto‑fit column after inserting line breaks | how to resize Excel column for multiline cell in Aspose | auto size column for wrapped content Aspose.Cells .NET | worksheet.AutoFitColumn example with text wrap
// Developer Intent: Automatically expand column A so every line of the wrapped multiline string in cell A1 is fully visible.
// Use Cases: Generating reports where description fields contain line breaks and need column auto‑sizing. | Designing invoice templates with product details that wrap and require dynamic column widths. | Exporting log entries with multi‑line messages to Excel while preserving readability.
// AI Prompts: Show C# code using Aspose.Cells to insert a multiline string, enable text wrapping, and auto‑fit the column for that cell. | Demonstrate how to call worksheet.AutoFitColumn for a specific column and row range after applying text wrap with newline characters. | Explain the parameters of AutoFitColumn and how they affect column width when the cell contains wrapped text.

using Aspose.Cells;
using System;

// Creates a workbook, inserts a newline‑separated string into cell A1, enables text wrapping, calls worksheet.AutoFitColumn to size column A to the wrapped content, and saves the file as AutoFitColumnMultiline.xlsx.
class AutoFitColumnExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert multiline text into cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("First line\nSecond line with more text\nThird line");

        // Enable text wrapping for the cell so the text occupies multiple lines
        Style style = cell.GetStyle();
        style.IsTextWrapped = true;
        cell.SetStyle(style);

        // Auto-fit column A (index 0) for the rows that contain the multiline text
        // Here we autofit only row 0 (A1), but you can specify a range of rows if needed
        worksheet.AutoFitColumn(0, 0, 0);

        // Save the workbook to a file
        workbook.Save("AutoFitColumnMultiline.xlsx");
    }
}

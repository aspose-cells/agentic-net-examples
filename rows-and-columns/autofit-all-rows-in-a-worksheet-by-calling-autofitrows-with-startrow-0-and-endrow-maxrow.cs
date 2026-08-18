// Title: C# – AutoFit all rows in an Aspose.Cells worksheet (startRow 0 to MaxDataRow)
// Description: Creates a Workbook, optionally adds sample text, determines the last populated row with MaxDataRow, calls Worksheet.AutoFitRows(0, maxRow) to resize every row to its content, and saves the file as AutoFitAllRows.xlsx.
// Keywords: Aspose.Cells AutoFitRows C# | auto fit rows Aspose .NET | adjust row height programmatically | MaxDataRow worksheet | Excel row auto‑size example | Aspose.Cells row height | C# Excel automation
// Common Searches: Aspose.Cells AutoFitRows all rows example | C# auto‑fit rows from first to last data row | How to resize rows automatically in Aspose.Cells | Get MaxDataRow and auto‑fit rows Aspose .NET | AutoFitRows startRow 0 endRow MaxDataRow
// Developer Intent: Resize every row that contains data so its height matches the cell content.
// Use Cases: Generating reports where wrapped text must be fully visible without manual formatting. | Exporting dynamic data to Excel and ensuring rows are sized correctly before distribution. | Building a template that adds rows programmatically and automatically adjusts their heights.
// AI Prompts: Show a C# snippet that auto‑fits rows from row 0 to the last populated row using Aspose.Cells, handling an empty worksheet gracefully. | Write a reusable method that accepts a Worksheet and applies AutoFitRows to all rows, including error handling for MaxDataRow = -1. | Explain how AutoFitRows works with wrapped text, merged cells, and custom row heights in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a Workbook, optionally adds sample text, determines the last populated row with MaxDataRow, calls Worksheet.AutoFitRows(0, maxRow) to resize every row to its content, and saves the file as AutoFitAllRows.xlsx.
class AutoFitAllRows
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data – this part can be omitted if the worksheet already has content
        sheet.Cells["A1"].PutValue("First row with a long text that should cause the row to expand when auto‑fitted.");
        sheet.Cells["A2"].PutValue("Second row");
        sheet.Cells["A3"].PutValue("Third row with\nmultiple lines\nto test auto‑fit.");

        // Determine the index of the last row that contains data (zero‑based)
        int maxRow = sheet.Cells.MaxDataRow;

        // Auto‑fit all rows from the first row (0) to the last data row
        sheet.AutoFitRows(0, maxRow);

        // Save the workbook
        workbook.Save("AutoFitAllRows.xlsx");
    }
}

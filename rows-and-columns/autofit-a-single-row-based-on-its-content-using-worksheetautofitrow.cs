// Title: How to auto‑fit the height of a single worksheet row using Worksheet.AutoFitRow in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells auto fit row height for a specific row | Worksheet.AutoFitRow usage example to adjust row size based on cell text | programmatically resize row height in a .NET workbook with Aspose.Cells | auto fit a row after inserting long text in Aspose.Cells C# | adjust Excel row height automatically for a single row using Aspose.Cells
// Tags: auto fit row height Aspose.Cells C# | Worksheet.AutoFitRow method example | resize single worksheet row based on content | adjust Excel row height programmatically .NET | long text row auto‑fit Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, writes a long string into cell A1, calls Worksheet.AutoFitRow(0) to automatically resize the first row's height according to the cell content, and saves the file as AutoFitRowResult.xlsx.
class AutoFitRowExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put some long text into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("This is a very long text that should cause the row height to increase when auto‑fitted.");

        // Auto‑fit the first row (row index is zero‑based)
        sheet.AutoFitRow(0);

        // Save the workbook (lifecycle save rule)
        workbook.Save("AutoFitRowResult.xlsx");
    }
}

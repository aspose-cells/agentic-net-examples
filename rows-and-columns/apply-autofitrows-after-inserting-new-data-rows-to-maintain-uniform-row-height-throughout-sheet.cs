// Title: Auto‑fit all rows after inserting new rows in an Aspose.Cells worksheet using C#
// AI Prompts: Insert rows at a specific index in a worksheet and then call worksheet.AutoFitRows to adjust every row height in C# with Aspose.Cells. | Create code that populates newly inserted rows and applies AutoFitRows before saving the workbook.
// Common Searches: aspocells c# auto fit rows after inserting rows | how to keep row heights uniform after adding rows in Excel with Aspose.Cells | c# insert rows and auto‑fit row height using Aspose.Cells | example of worksheet.AutoFitRows after InsertRows in C#
// Tags: auto‑fit rows after row insertion Aspose.Cells | insert rows with worksheet.AutoFitRows C# | adjust row height after adding rows Aspose.Cells | Aspose.Cells AutoFitRows usage C# | maintain consistent row height in Excel Aspose.Cells

using System;
using Aspose.Cells;

// Shows how to create a workbook, insert rows at a given index, fill the new rows with data, call worksheet.AutoFitRows to ensure uniform row heights, and save the file.
class AutoFitRowsAfterInsertDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add some initial data
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Original Row 1");
        cells["A3"].PutValue("Original Row 2");

        // Insert two new rows at index 2 (between the original rows)
        // Uses the InsertRows(int rowIndex, int totalRows) rule
        cells.InsertRows(2, 2);

        // Populate the newly inserted rows
        cells["A3"].PutValue("Inserted Row 1");
        cells["A4"].PutValue("Inserted Row 2");

        // Auto‑fit all rows in the worksheet to ensure uniform row height
        // Uses the AutoFitRows() rule
        worksheet.AutoFitRows();

        // Save the workbook (lifecycle rule)
        workbook.Save("AutoFitRowsAfterInsert.xlsx");
    }
}

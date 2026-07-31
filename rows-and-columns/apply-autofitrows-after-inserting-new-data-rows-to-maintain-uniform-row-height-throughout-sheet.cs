// Title: AutoFitRows after inserting rows with Aspose.Cells for .NET (C#)
// Description: This example shows how to create a workbook, insert new rows, fill them with data, and call worksheet.AutoFitRows() to automatically adjust every row's height so the sheet keeps a consistent appearance before saving.
// Keywords: Aspose.Cells AutoFitRows | C# AutoFitRows after insert | worksheet.AutoFitRows usage | adjust Excel row height programmatically | insert rows Aspose.Cells | auto fit rows .NET | Excel row height automation
// Common Searches: Aspose.Cells AutoFitRows after inserting rows | C# auto fit rows in Excel workbook | how to adjust row height after adding rows Aspose.Cells | worksheet.AutoFitRows example C# | auto‑fit all rows after bulk insert Aspose.Cells
// Developer Intent: Resize all rows automatically after new rows are added to maintain uniform height.
// Use Cases: Add a variable number of data rows to a template and auto‑fit them before exporting. | Generate dynamic reports where row count changes at runtime and consistent layout is required. | Replace placeholder rows in an existing sheet with actual content and keep row heights consistent.
// AI Prompts: Generate C# code that inserts rows at a given index and then calls worksheet.AutoFitRows() using Aspose.Cells. | Explain the performance impact of calling AutoFitRows after bulk row insertion in a large workbook. | Show how to combine InsertRows and AutoFitRows to maintain a tidy Excel sheet in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitRowsDemo
{
    // This example shows how to create a workbook, insert new rows, fill them with data, and call worksheet.AutoFitRows() to automatically adjust every row's height so the sheet keeps a consistent appearance before saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some initial data
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Row 1 data");
            cells["A3"].PutValue("Row 2 data");
            cells["A4"].PutValue("Row 3 data");

            // Insert two new rows after the header (at index 1)
            // This pushes existing rows down and creates empty rows ready for new data
            cells.InsertRows(1, 2);

            // Add data to the newly inserted rows
            cells["A2"].PutValue("Inserted Row 1");
            cells["A3"].PutValue("Inserted Row 2");

            // Apply AutoFitRows to adjust the height of all rows based on their content
            // This ensures uniform row height throughout the sheet
            worksheet.AutoFitRows();

            // Save the workbook to a file
            workbook.Save("AutoFitRowsAfterInsert.xlsx");
        }
    }
}

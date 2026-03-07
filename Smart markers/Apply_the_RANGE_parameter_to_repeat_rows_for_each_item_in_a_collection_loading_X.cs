using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook (template)
        Workbook workbook = new Workbook("Template.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Row that contains the template (zero‑based index)
        int templateRowIndex = 1; // e.g., second row in the sheet

        // Collection of data items – each item will produce a new row
        List<string[]> dataItems = new List<string[]>
        {
            new[] { "John", "Doe", "30" },
            new[] { "Jane", "Smith", "25" },
            new[] { "Bob", "Brown", "40" }
        };

        // Position where the first copied row will be inserted
        int insertRowIndex = templateRowIndex + 1;

        foreach (var item in dataItems)
        {
            // Insert an empty row at the current position
            cells.InsertRow(insertRowIndex);

            // Copy the template row into the newly inserted row
            cells.CopyRows(cells, templateRowIndex, insertRowIndex, 1);

            // Populate the copied row with the current item's values
            for (int col = 0; col < item.Length; col++)
            {
                cells[insertRowIndex, col].PutValue(item[col]);
            }

            // Advance the insert position for the next iteration
            insertRowIndex++;
        }

        // If the original template row is no longer needed, delete it:
        // cells.DeleteRow(templateRowIndex);

        // Save the workbook with the repeated rows
        workbook.Save("Result.xlsx");
    }
}
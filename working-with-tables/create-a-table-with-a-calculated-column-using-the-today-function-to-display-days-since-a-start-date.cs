// Title: Add a TODAY() calculated column to an Excel table using Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, builds a ListObject covering A1:C4, and assigns the structured‑reference formula "=TODAY()-[@StartDate]" to the DaysSinceStart column. The code calculates the formulas and saves the file as TableWithCalculatedColumn.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel table | ListObject | calculated column | TODAY function | structured reference | formula assignment | date difference
// Common Searches: Aspose.Cells add calculated column TODAY | C# set formula for Excel table column Aspose | How to use structured references in Aspose.Cells | Compute days since date in Excel with Aspose.Cells | Aspose.Cells ListObject formula example
// Developer Intent: Create an Excel ListObject in Aspose.Cells and populate a column with a TODAY()‑based formula that returns the number of days elapsed since each start date.
// Use Cases: Project tracker that automatically shows days elapsed from a start date. | Order‑aging report that calculates days since order placement. | Dashboard highlighting overdue tasks by computing days since a reference date.
// AI Prompts: Show C# code to add a TODAY()-based calculated column to an Aspose.Cells table using structured references. | How can I force formula recalculation after setting a column formula in Aspose.Cells? | Explain how to format the DaysSinceStart column as a number and handle negative values in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This example creates a new workbook, builds a ListObject covering A1:C4, and assigns the structured‑reference formula "=TODAY()-[@StartDate]" to the DaysSinceStart column. The code calculates the formulas and saves the file as TableWithCalculatedColumn.xlsx.
class TableWithCalculatedColumn
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ----- Fill data for the table -----
        // Header row
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("StartDate");
        cells["C1"].PutValue("DaysSinceStart");

        // Sample data rows
        cells["A2"].PutValue(1);
        cells["B2"].PutValue(new DateTime(2023, 1, 1));

        cells["A3"].PutValue(2);
        cells["B3"].PutValue(new DateTime(2023, 2, 15));

        cells["A4"].PutValue(3);
        cells["B4"].PutValue(new DateTime(2023, 3, 10));

        // ----- Create a ListObject (table) covering the data range -----
        // The range includes the header row and the three data rows (A1:C4)
        int tableIndex = sheet.ListObjects.Add("A1", "C4", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // ----- Add a calculated column "DaysSinceStart" -----
        // The column already exists (C). We'll set a formula for each data row.
        // Structured reference uses the @ symbol to refer to the current row.
        string formula = "=TODAY()-[@StartDate]";

        // Number of data rows in the table (excluding header)
        int dataRowCount = table.DataRange.RowCount; // rows after header

        // Row offset is zero‑based relative to the first data row.
        for (int i = 0; i < dataRowCount; i++)
        {
            // Column offset 2 corresponds to the third column (C) in the table.
            table.PutCellFormula(i, 2, formula);
        }

        // Calculate all formulas so the "DaysSinceStart" values are populated
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("TableWithCalculatedColumn.xlsx");
    }
}

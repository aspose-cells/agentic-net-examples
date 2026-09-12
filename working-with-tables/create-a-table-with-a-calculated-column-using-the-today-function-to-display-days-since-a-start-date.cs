// Title: Add a calculated column that shows days since a start date using TODAY() in an Aspose.Cells Excel table (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a worksheet, inserts a ListObject table, and sets the third column formula to =TODAY()-[StartDate]. | Rename the start‑date column to "BeginDate" and adjust the calculated column formula to reference the new header. | Apply a medium‑style table format, force formula recalculation, and save the workbook as TableWithCalculatedColumn.xlsx.
// Common Searches: how to use TODAY() in a calculated column of an Aspose.Cells ListObject table in C# | Aspose.Cells C# create Excel table with days‑since‑start column | C# Aspose.Cells structured reference formula for date difference in a table | calculate days between today and a date column using Aspose.Cells | apply table style and recalculate formulas in Aspose.Cells workbook C#
// Tags: Aspose.Cells calculated column TODAY() formula | C# ListObject table structured reference | Excel table style assignment Aspose.Cells | recalculate workbook formulas Aspose.Cells | save workbook as .xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program creates a new workbook, adds a ListObject table with ID, StartDate, and DaysSinceStart columns, defines a calculated column using the formula =TODAY()-[StartDate], applies a medium table style, forces formula recalculation, and saves the file as TableWithCalculatedColumn.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add headers
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("StartDate");
            sheet.Cells["C1"].PutValue("DaysSinceStart");

            // Add sample data
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(new DateTime(2023, 6, 15));
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(new DateTime(2024, 2, 20));

            // Determine the range that will become the table (including header)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = sheet.Cells.MaxDataRow + 1; // includes header row
            int totalColumns = 3;      // ID, StartDate, DaysSinceStart

            // Add a ListObject (Excel table) to the worksheet; hasHeaders = true
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            // Table name and header visibility are already set by the Add method
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Set a calculated column formula: DaysSinceStart = TODAY() - StartDate
            // Use structured reference to the StartDate column
            ListColumn calcColumn = table.ListColumns[2]; // third column (zero‑based)
            calcColumn.Formula = "=TODAY()-[StartDate]";

            // Recalculate formulas so the workbook contains the computed values
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("TableWithCalculatedColumn.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

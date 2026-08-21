// Title: Aspose.Cells .NET: Add a Calculated Column with TODAY() to Show Days Since Start Date
// Description: C# example that creates a new workbook, defines an Excel table (ListObject) with "StartDate" and "DaysSinceStart" columns, inserts sample dates, applies the formula "=TODAY()-[@StartDate]" to the calculated column, forces formula evaluation, and saves the file as TableWithDaysSinceStart.xlsx.
// Keywords: Aspose.Cells | .NET | C# | calculated column | TODAY() function | Excel table formula | ListObject | days since date | auto recalc | save workbook
// Common Searches: Aspose.Cells add calculated column to table | use TODAY() in Aspose.Cells formula | compute days elapsed from date column .NET | ListObject formula example Aspose.Cells | auto‑recalculate formulas Aspose.Cells C#
// Developer Intent: Create an Excel table with a calculated column that returns the number of days elapsed from each start date using the TODAY() function.
// Use Cases: Project status reports that automatically display days since project kickoff. | Inventory aging dashboards that calculate days since receipt. | Support ticket aging analysis showing how many days each ticket has been open.
// AI Prompts: Generate C# code using Aspose.Cells to build a table and add a calculated column that uses TODAY() to compute days since a date field. | Show how to assign a formula to every row of a ListObject's calculated column and trigger workbook recalculation in Aspose.Cells. | Explain how to format the calculated column as a number and save the workbook after applying the TODAY()‑based formula.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// C# example that creates a new workbook, defines an Excel table (ListObject) with "StartDate" and "DaysSinceStart" columns, inserts sample dates, applies the formula "=TODAY()-[@StartDate]" to the calculated column, forces formula evaluation, and saves the file as TableWithDaysSinceStart.xlsx.
class TableWithCalculatedColumn
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers for the table
            sheet.Cells["A1"].PutValue("StartDate");
            sheet.Cells["B1"].PutValue("DaysSinceStart");

            // Populate some start dates
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 3, 15));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 6, 30));
            sheet.Cells["A5"].PutValue(new DateTime(2023, 9, 10));

            // Create a ListObject (Excel table) covering the range A1:B5
            int tableIdx = sheet.ListObjects.Add("A1", "B5", true);
            ListObject table = sheet.ListObjects[tableIdx];

            // Formula for the calculated column: TODAY() - [StartDate]
            string formula = "=TODAY()-[@StartDate]";

            // Apply the formula to each data row in the "DaysSinceStart" column
            // Row offset 0 is the header; data rows start at 1.
            int dataRows = table.DataRange.RowCount; // number of data rows
            for (int i = 1; i <= dataRows; i++)
            {
                table.PutCellFormula(i, 1, formula);
            }

            // Recalculate formulas so the calculated column is populated
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("TableWithDaysSinceStart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

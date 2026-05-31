using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class TableWithCalculatedColumnDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define headers for the table
                cells["A1"].PutValue("StartDate");      // Column with the start date
                cells["B1"].PutValue("DaysSinceStart"); // Calculated column

                // Populate some start dates (example dates)
                cells["A2"].PutValue(new DateTime(2023, 1, 1));
                cells["A3"].PutValue(new DateTime(2023, 2, 15));
                cells["A4"].PutValue(new DateTime(2023, 3, 10));
                cells["A5"].PutValue(new DateTime(2023, 4, 20));

                // Create a ListObject (Excel table) that includes the data range A1:B5
                // The last parameter 'true' indicates that the first row contains headers
                int tableIndex = sheet.ListObjects.Add("A1", "B5", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Set the formula for the calculated column using a structured reference.
                // The formula calculates the number of days between TODAY() and the start date in the same row.
                // Row offset starts at 1 for the first data row (row 2 in the worksheet).
                for (int rowOffset = 1; rowOffset <= 4; rowOffset++)
                {
                    // Column offset 1 corresponds to the second column ("DaysSinceStart")
                    table.PutCellFormula(rowOffset, 1, "=TODAY()-[@StartDate]");
                }

                // Recalculate all formulas so that the new column shows the correct values
                workbook.CalculateFormula();

                // Save the workbook
                workbook.Save("TableWithCalculatedColumn.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TableWithCalculatedColumnDemo.Run();
        }
    }
}
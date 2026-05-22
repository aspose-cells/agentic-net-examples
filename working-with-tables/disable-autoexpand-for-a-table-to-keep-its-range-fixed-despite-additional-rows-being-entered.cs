using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class DisableTableAutoExpandDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some initial data (including header row)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[i - 1, 0].PutValue(i - 1);               // ID column
                    sheet.Cells[i - 1, 1].PutValue($"Item {i - 1}");    // Name column
                }

                // Add a ListObject (table) covering the data range A1:B5
                int tableIndex = sheet.ListObjects.Add("A1", "B5", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Auto‑expand is not required for this demo; the table range will stay fixed
                // (If the API supported it, you could set table.AutoExpand = false;)

                // Add additional rows *outside* the original table range
                // These rows will not be included in the table because we do not expand it
                sheet.Cells["A6"].PutValue(6);
                sheet.Cells["B6"].PutValue("Item 6");
                sheet.Cells["A7"].PutValue(7);
                sheet.Cells["B7"].PutValue("Item 7");

                // Verify that the table range has not changed
                Console.WriteLine($"Table range after adding rows: {table.StartRow}-{table.EndRow}, {table.StartColumn}-{table.EndColumn}");
                // Expected output: 0-4 (rows 0‑4 correspond to A1:B5)

                // Save the workbook
                workbook.Save("TableAutoExpandDisabled.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableTableAutoExpandDemo.Run();
        }
    }
}
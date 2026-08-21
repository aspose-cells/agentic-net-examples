// Title: Aspose.Cells for .NET – Hide a ListObject totals row (C#)
// Description: Demonstrates how to create a workbook, add a ListObject covering A1:B4, and remove its totals row by setting the ShowTotals property to false, then save the file as RemoveTableTotalsRowDemo.xlsx.
// Keywords: Aspose.Cells C# hide totals row | ListObject ShowTotals false | remove table totals row Aspose.Cells | Excel table hide totals row .NET | Aspose.Cells ListObject properties | C# programmatically hide Excel table totals
// Common Searches: how to hide totals row in Aspose.Cells | Aspose.Cells set ShowTotals false | remove totals row from Excel table C# | Aspose.Cells ListObject hide totals | C# hide table totals row Aspose
// Developer Intent: Disable the display of the totals row for a worksheet table by setting ListObject.ShowTotals to false.
// Use Cases: Generate reports where a totals row is not required. | Provide users an option to toggle the totals row when exporting data. | Create clean data tables for downstream processing without aggregate rows.
// AI Prompts: Write C# code using Aspose.Cells to add a ListObject and hide its totals row. | Explain the effect of the ShowTotals property on an Excel table in Aspose.Cells. | Show how to conditionally hide or show a table's totals row based on a boolean flag in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a ListObject covering A1:B4, and remove its totals row by setting the ShowTotals property to false, then save the file as RemoveTableTotalsRowDemo.xlsx.
    public class RemoveTableTotalsRowDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["B3"].PutValue(15);
                worksheet.Cells["A4"].PutValue("Banana");
                worksheet.Cells["B4"].PutValue(8);

                // Add a ListObject (table) covering the data range
                int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Ensure the totals row is initially visible (optional)
                table.ShowTotals = true;

                // Hide the totals row as required
                table.ShowTotals = false;

                // Save the workbook
                workbook.Save("RemoveTableTotalsRowDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}

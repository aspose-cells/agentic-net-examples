// Title: Capture the exception thrown when adding an AutoFilter after converting a ListObject to a regular range using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a ListObject on a worksheet, calls ConvertToRange, then attempts to add an AutoFilter on a column and writes the caught exception type and message to the console. | Modify the example to apply the AutoFilter on a different column after the ListObject conversion and log the resulting exception details while still saving the workbook.
// Common Searches: Aspose.Cells exception when applying AutoFilter after ListObject.ConvertToRange | how to catch filter error after converting Excel table to range in C# | Add filter to removed table using Aspose.Cells throws what exception | C# Aspose.Cells AutoFilter on converted ListObject example | handling AutoFilter failure after Excel table conversion with Aspose.Cells
// Tags: ListObject ConvertToRange Aspose.Cells C# | AutoFilter exception after table conversion | capture Aspose.Cells filter error | Excel table to range conversion handling | Aspose.Cells remove ListObject filter operation

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates creating a ListObject, converting it to a normal range, attempting to add an AutoFilter on the former table, catching the resulting exception, and saving the workbook.
class ApplyFilterAfterTableConversion
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1:C5)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue("Diana");
            sheet.Cells["C5"].PutValue(88);

            // Add a table (ListObject) over the data range A1:C5
            int firstRow = 0;   // zero‑based index
            int firstColumn = 0;
            int totalRows = 5;
            int totalColumns = 3;
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Convert the table back to a normal range (removes the ListObject)
            table.ConvertToRange();

            // Attempt to apply an AutoFilter to the former table range.
            // Since the ListObject has been removed, accessing its AutoFilter will raise an exception.
            try
            {
                table.AutoFilter.AddFilter(1, "Alice");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Expected exception caught:");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"Type: {ex.GetType().FullName}");
            }

            // Save the workbook (no filter applied)
            workbook.Save("Result.xlsx");
        }
        catch (Exception e)
        {
            Console.WriteLine("An unexpected error occurred:");
            Console.WriteLine(e.Message);
        }
    }
}

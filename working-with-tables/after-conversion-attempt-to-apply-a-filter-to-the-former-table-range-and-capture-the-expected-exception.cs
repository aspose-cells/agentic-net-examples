// Title: Capture the exception thrown by AutoFilter after converting a ListObject to a range with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# AutoFilter throws exception after ListObject.ConvertToRange | catch exception when applying filter to a converted table in Aspose.Cells | listobject to range then use AutoFilter error handling in .NET | expected exception type from AutoFilter on non‑table range Aspose.Cells | how to retrieve exception details from AutoFilter after table conversion in C#
// Tags: ListObject.ConvertToRange exception handling | AutoFilter usage after table conversion | Aspose.Cells C# filter error on range | capturing Aspose.Cells AutoFilter exception | working with tables and ranges Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds a ListObject covering a data range, converts the ListObject back to a normal range, and then attempts to apply an AutoFilter. Because the table no longer exists, AutoFilter throws an exception, which is caught and its type and message are displayed before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (A1:C5)
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

            // Add a table (ListObject) covering the data range A1:C5
            int firstRow = 0;        // zero‑based index for row 1
            int firstColumn = 0;     // zero‑based index for column A
            int totalRows = 5;
            int totalColumns = 3;

            // Add returns the index of the newly created ListObject
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Convert the table back to a normal range
            table.ConvertToRange();

            // After conversion, attempting to use the ListObject's AutoFilter should raise an exception
            try
            {
                // This line is expected to throw because the ListObject no longer exists as a table
                table.AutoFilter.AddFilter(0, "1");
            }
            catch (Exception ex)
            {
                // Capture and display the expected exception details
                Console.WriteLine("Expected exception caught:");
                Console.WriteLine($"Type: {ex.GetType().FullName}");
                Console.WriteLine($"Message: {ex.Message}");
            }

            // Save the workbook (ensure the directory is writable)
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception e)
        {
            // General exception handling to avoid unexpected crashes
            Console.WriteLine("An error occurred:");
            Console.WriteLine(e.ToString());
        }
    }
}

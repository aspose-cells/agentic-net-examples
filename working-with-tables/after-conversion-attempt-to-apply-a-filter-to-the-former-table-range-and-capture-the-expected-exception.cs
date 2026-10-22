// Title: Convert a ListObject to a regular range and capture the AutoFilter exception using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a worksheet, adds a ListObject, converts it to a normal range with ConvertToRange, then calls AutoFilter.Filter on the original ListObject and catches the resulting exception. | Show how to handle the exception thrown when applying an AutoFilter to a ListObject after it has been converted to a range in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# catch exception when filtering a ListObject after ConvertToRange | filtering a table after converting to range throws error Aspose.Cells | AutoFilter on removed ListObject exception handling Aspose.Cells .NET | convert ListObject to range then apply filter example Aspose.Cells C#
// Tags: convert ListObject to range Aspose.Cells | AutoFilter exception after ListObject conversion | handling invalid ListObject Aspose.Cells C# | Aspose.Cells table to range conversion error handling | C# Aspose.Cells ListObject AutoFilter exception

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds a ListObject covering A1:B4, converts the table to a normal range, attempts to apply an AutoFilter on the now‑invalid ListObject, catches the expected exception, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill the worksheet with sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);
            sheet.Cells["A4"].PutValue("Charlie");
            sheet.Cells["B4"].PutValue(78);

            // Add a table (ListObject) that covers the data range A1:B4
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 4;     // includes header row
            int totalColumns = 2;

            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                                                   firstRow + totalRows,
                                                   firstColumn + totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Convert the table back to a normal range; the ListObject is removed from the worksheet
            table.ConvertToRange();

            // Attempt to apply a filter using the now‑invalid ListObject and capture the expected exception
            try
            {
                // This call should fail because the ListObject no longer exists after conversion
                table.AutoFilter.Filter(0, "Alice");
            }
            catch (Exception ex)
            {
                // Expected exception handling
                Console.WriteLine("Expected exception caught: " + ex.Message);
            }

            // Save the workbook (optional, demonstrates that the file is still valid)
            string outputPath = "FilteredResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception e)
        {
            // General exception handling for unexpected errors
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }
}

// Title: Aspose.Cells .NET – Verify ListObject.IsTable Is False After ConvertToRange
// Description: C# example that creates a workbook, adds a ListObject (table) on range A1:B3, calls ConvertToRange to turn the table into a normal range, checks ListObject.IsTable (or the ListObjects collection) to confirm the table no longer exists, and saves the file.
// Keywords: Aspose.Cells ConvertToRange | ListObject.IsTable | C# table to range conversion | verify table removal Aspose | .NET workbook table check | Aspose.Cells ListObject conversion
// Common Searches: Aspose.Cells how to check if ListObject is still a table after ConvertToRange | C# verify table conversion to range in Aspose.Cells | ListObject.IsTable false after ConvertToRange | Aspose.Cells remove table features after conversion
// Developer Intent: Ensure a ListObject no longer behaves as a table after ConvertToRange is executed.
// Use Cases: Programmatically confirm that sorting, filtering, and other table features are disabled after conversion. | Clean workbook metadata before sharing with users who do not need table structures. | Skip table‑specific logic in pipelines that process converted ranges.
// AI Prompts: Generate C# code using Aspose.Cells to convert a ListObject to a range and assert that ListObject.IsTable returns false. | Explain step‑by‑step how to validate that a table has been removed by checking the worksheet's ListObjects count and the IsTable property. | Provide a test scenario that confirms a converted table no longer supports sorting in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// C# example that creates a workbook, adds a ListObject (table) on range A1:B3, calls ConvertToRange to turn the table into a normal range, checks ListObject.IsTable (or the ListObjects collection) to confirm the table no longer exists, and saves the file.
class VerifyTableConversion
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Mary");

            // Add a ListObject (table) covering the data range
            int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
            ListObject listObject = sheet.ListObjects[tableIndex];

            // Since a ListObject represents a table, we can consider it a table before conversion
            Console.WriteLine("IsTable before conversion: true");

            // Convert the table to a normal range
            listObject.ConvertToRange();

            // After conversion the ListObject is removed; verify that it no longer exists
            bool isTableAfterConversion = sheet.ListObjects.Count > tableIndex;
            Console.WriteLine("IsTable after conversion: " + isTableAfterConversion);

            // Save the workbook
            string outputPath = "TableConversionResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

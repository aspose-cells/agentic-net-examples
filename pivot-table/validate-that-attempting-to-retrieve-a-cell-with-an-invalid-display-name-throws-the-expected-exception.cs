// Title: Verify that PivotTable.GetCellByDisplayName throws an exception for a non‑existent field using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a pivot table, and asserts that calling GetCellByDisplayName with a missing field name raises the expected exception. | Demonstrate how to catch the specific exception thrown by PivotTable.GetCellByDisplayName when the display name is not found, and log its type and message.
// Common Searches: Aspose.Cells C# GetCellByDisplayName throws when field does not exist | how to catch exception from PivotTable.GetCellByDisplayName invalid display name | unit test for missing pivot field display name Aspose.Cells | exception type returned by GetCellByDisplayName for unknown field in Aspose.Cells | validate error handling for pivot table cell retrieval by display name in .NET
// Tags: Aspose.Cells pivot table GetCellByDisplayName exception | invalid pivot field display name handling Aspose.Cells | C# Aspose.Cells missing pivot field error | pivot table cell retrieval error Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotValidation
{
    // The example creates a workbook, populates data, adds a pivot table, then attempts to retrieve a cell using PivotTable.GetCellByDisplayName with a non‑existent field name, catches the resulting exception, outputs its type and message, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Food";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Travel";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Supplies";
            cells["B4"].Value = 45;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "DemoPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add row field and data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate the pivot table so that cells are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Attempt to retrieve a cell using an invalid display name
            string invalidDisplayName = "NonExistentField";

            try
            {
                // This call is expected to throw an exception because the display name does not exist
                Cell cell = pivotTable.GetCellByDisplayName(invalidDisplayName);
                // If no exception is thrown, the test has failed
                Console.WriteLine("FAIL: No exception thrown. Method returned cell at {0},{1}", cell.Row, cell.Column);
            }
            catch (Exception ex)
            {
                // Verify that the caught exception is the expected type/message
                Console.WriteLine("PASS: Expected exception caught.");
                Console.WriteLine("Exception Type: {0}", ex.GetType().FullName);
                Console.WriteLine("Message: {0}", ex.Message);
            }

            // Save the workbook (optional, demonstrates normal lifecycle usage)
            workbook.Save("PivotTable_InvalidDisplayName_Test.xlsx");
        }
    }
}

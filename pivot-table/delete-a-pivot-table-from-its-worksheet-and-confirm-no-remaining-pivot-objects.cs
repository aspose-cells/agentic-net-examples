// Title: How to delete a PivotTable from a worksheet and confirm its removal using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that deletes a specific PivotTable from a worksheet and verifies that the worksheet's PivotTables collection is empty using Aspose.Cells. | Show an example of removing a PivotTable, then printing the PivotTables count before and after deletion with Aspose.Cells for .NET. | Demonstrate how to call the Remove method on Worksheet.PivotTables and check for remaining pivot objects in a C# Aspose.Cells workbook.
// Common Searches: c# aspnet delete pivot table from worksheet using Aspose.Cells | how to check if a PivotTable was removed in Aspose.Cells .NET | Aspose.Cells remove specific pivot table and get pivot tables count | sample code for deleting a pivot table with Aspose.Cells for C# | verify pivot table deletion in Aspose.Cells workbook
// Tags: Aspose.Cells remove pivot table | Aspose.Cells pivot table count | Aspose.Cells delete pivot table worksheet | Aspose.Cells refresh pivot data | Aspose.Cells C# pivot table lifecycle

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a PivotTable, refreshes and calculates it, then removes the PivotTable using Worksheet.PivotTables.Remove, prints the PivotTables count before and after removal, and saves the file.
    public class DeletePivotTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Apple");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(130);

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh data and calculate the pivot table
                pivotTable.RefreshData();      // Correct API to refresh the pivot cache
                pivotTable.CalculateData();

                // Verify that the pivot table exists
                Console.WriteLine("Pivot tables count before removal: " + sheet.PivotTables.Count);

                // Remove the pivot table using the Remove method
                sheet.PivotTables.Remove(pivotTable);

                // Confirm that no pivot tables remain in the worksheet
                Console.WriteLine("Pivot tables count after removal: " + sheet.PivotTables.Count);

                // Save the workbook (optional, just to demonstrate lifecycle compliance)
                workbook.Save("PivotTableDeleted.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DeletePivotTableDemo.Run();
        }
    }
}

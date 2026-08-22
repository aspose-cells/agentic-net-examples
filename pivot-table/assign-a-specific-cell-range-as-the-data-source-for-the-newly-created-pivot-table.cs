// Title: How to reassign a pivot table’s data source to a different cell range using Aspose.Cells for .NET (C#)
// AI Prompts: Create a pivot table from range A1:B5, then switch its data source to C1:D5 and refresh it with Aspose.Cells in C#. | Programmatically change the source range of an existing Aspose.Cells pivot table and recalculate the pivot cache using the ChangeDataSource method. | Demonstrate updating a pivot table’s data source to a new worksheet range and saving the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# change pivot table source range from A1:B5 to C1:D5 | How to use ChangeDataSource method to update pivot table data source in Aspose.Cells .NET | Refresh pivot cache after changing data source in Aspose.Cells C# example | Assign new data range to existing pivot table using Aspose.Cells for .NET
// Tags: Aspose.Cells pivot table source update | C# modify pivot table data range | refresh pivot cache Aspose.Cells | Aspose.Cells workbook pivot table example | pivot table data source reassignment .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // The sample creates a workbook, fills A1:B5 with sample data, adds a pivot table at D3, copies similar data to C1:D5, changes the pivot table's data source to the new range using ChangeDataSource, refreshes and recalculates the pivot, and saves the file as AssignDataSourceDemo.xlsx.
    public class AssignDataSourceDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in the worksheet (range A1:B5)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("C");
                sheet.Cells["B5"].PutValue(40);

                // Add a pivot table using the initial data range A1:B5
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "MyPivotTable");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table (add fields)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Define a new data source range (C1:D5) on the same worksheet
                // For demonstration, copy the original data to the new range
                sheet.Cells["C1"].PutValue("Category");
                sheet.Cells["D1"].PutValue("Value");
                sheet.Cells["C2"].PutValue("X");
                sheet.Cells["D2"].PutValue(100);
                sheet.Cells["C3"].PutValue("Y");
                sheet.Cells["D3"].PutValue(200);
                sheet.Cells["C4"].PutValue("X");
                sheet.Cells["D4"].PutValue(300);
                sheet.Cells["C5"].PutValue("Z");
                sheet.Cells["D5"].PutValue(400);

                // Change the pivot table's data source to the new range C1:D5
                string[] newDataSource = new string[] { "C1:D5" };
                pivotTable.ChangeDataSource(newDataSource);

                // Refresh the pivot cache and recalculate the pivot table to reflect the new source
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "AssignDataSourceDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AssignDataSourceDemo.Run();
        }
    }
}

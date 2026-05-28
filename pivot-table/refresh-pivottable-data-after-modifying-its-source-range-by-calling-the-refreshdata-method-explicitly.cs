using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace Demo
{
    public class RefreshPivotTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(150);

                // Add a pivot table based on the source range A1:B4, placed at E3
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Product as row field, Sales as data field
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                pivot.AddFieldToArea(PivotFieldType.Data, 1);

                // Initial calculation of the pivot table
                pivot.CalculateData();

                // Modify the source data after the initial calculation
                sheet.Cells["B2"].PutValue(120); // Updated sales for product A
                sheet.Cells["B3"].PutValue(250); // Updated sales for product B

                // Refresh the pivot cache to reflect the changed source data
                pivot.RefreshData();

                // Recalculate the pivot table after refreshing the data
                pivot.CalculateData();

                // Save the workbook with the refreshed pivot table
                workbook.Save("RefreshPivotTableDemo.xlsx");
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
            RefreshPivotTableDemo.Run();
        }
    }
}
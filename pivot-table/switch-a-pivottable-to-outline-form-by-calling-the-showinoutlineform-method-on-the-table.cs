using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableOutlineFormDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet for source data
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data
                dataSheet.Cells["A1"].PutValue("Date");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Sales");

                dataSheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                dataSheet.Cells["B2"].PutValue("Apple");
                dataSheet.Cells["C2"].PutValue(1000);

                dataSheet.Cells["A3"].PutValue(new DateTime(2023, 1, 2));
                dataSheet.Cells["B3"].PutValue("Banana");
                dataSheet.Cells["C3"].PutValue(2000);

                dataSheet.Cells["A4"].PutValue(new DateTime(2023, 1, 3));
                dataSheet.Cells["B4"].PutValue("Apple");
                dataSheet.Cells["C4"].PutValue(1500);

                // Add a worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Create the pivot table and obtain its reference
                int pivotIndex = pivotSheet.PivotTables.Add(
                    "=Data!A1:C4",   // source data range
                    "A3",            // top‑left cell of the pivot table
                    "PivotTable1");  // pivot table name

                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add fields to the pivot table areas
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Switch the pivot table layout to Outline form
                pivotTable.ShowInOutlineForm();

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PivotTableOutlineFormDemo.xlsx");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
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
            PivotTableOutlineFormDemo.Run();
        }
    }
}
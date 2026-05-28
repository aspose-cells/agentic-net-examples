using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class DisableColumnGrandTotalsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Product";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Electronics";
                sheet.Cells["B2"].Value = "Phone";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "Electronics";
                sheet.Cells["B3"].Value = "Laptop";
                sheet.Cells["C3"].Value = 2500;

                sheet.Cells["A4"].Value = "Furniture";
                sheet.Cells["B4"].Value = "Chair";
                sheet.Cells["C4"].Value = 300;

                sheet.Cells["A5"].Value = "Furniture";
                sheet.Cells["B5"].Value = "Table";
                sheet.Cells["C5"].Value = 800;

                // Add a pivot table that uses the data range A1:C5 and place it at E3
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Product as column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

                // Disable column grand totals
                pivotTable.ShowColumnGrandTotals = false;

                // Save the workbook to a file
                string outputPath = "PivotTable_NoColumnGrandTotals.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DisableColumnGrandTotalsDemo.Run();
        }
    }
}
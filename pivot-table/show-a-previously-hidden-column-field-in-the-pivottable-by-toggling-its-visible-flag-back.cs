// Title: C# – Unhide a PivotTable Column Field with Aspose.Cells (ShowAllItems)
// Description: Creates a workbook, builds a PivotTable, then makes a hidden column field visible by setting ShowAllItems = true, refreshes the cache, recalculates data, and saves the file.
// Keywords: Aspose.Cells C# PivotTable hide column | unhide pivot column field Aspose | ShowAllItems property | PivotField visibility .NET | refresh pivot data Aspose.Cells | calculate pivot Aspose.Cells | C# Excel pivot programmatically
// Common Searches: Aspose.Cells unhide pivot column C# | ShowAllItems column field Aspose.Cells | make hidden pivot field visible .NET | refresh pivot after changing visibility Aspose | C# code to show all items in pivot column
// Developer Intent: Programmatically reveal a hidden column field in an Aspose.Cells PivotTable.
// Use Cases: Ensure every region appears as a column after creating a PivotTable by setting the first ColumnField's ShowAllItems to true. | Toggle column visibility based on user selection by updating the PivotField's ShowAllItems (or Visible) flag and then refreshing the pivot cache. | Prepare a report workbook where hidden pivot columns are automatically shown before exporting to Excel.
// AI Prompts: Generate C# code using Aspose.Cells to unhide a hidden column field in an existing PivotTable and refresh the data. | Show how to set the Visible/ShowAllItems property of a PivotField, recalculate the pivot, and save the workbook. | Explain the steps to programmatically display all items of a PivotTable column field with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, builds a PivotTable, then makes a hidden column field visible by setting ShowAllItems = true, refreshes the cache, recalculates data, and saves the file.
    public class ShowHiddenColumnFieldDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Region";
                cells["C1"].Value = "Sales";

                cells["A2"].Value = "Bike";
                cells["B2"].Value = "North";
                cells["C2"].Value = 1200;

                cells["A3"].Value = "Bike";
                cells["B3"].Value = "South";
                cells["C3"].Value = 1500;

                cells["A4"].Value = "Car";
                cells["B4"].Value = "North";
                cells["C4"].Value = 2000;

                cells["A5"].Value = "Car";
                cells["B5"].Value = "South";
                cells["C5"].Value = 2500;

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add a row field (Product) and a column field (Region)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");

                // Add a data field (Sales) and set its aggregation function
                int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.DataFields[dataFieldPos].Function = ConsolidationFunction.Sum;

                // Ensure the column field "Region" is visible
                if (pivotTable.ColumnFields.Count > 0)
                {
                    PivotField columnField = pivotTable.ColumnFields[0]; // "Region"
                    columnField.ShowAllItems = true; // Make all items visible
                }

                // Refresh pivot cache and recalculate data
                pivotTable.RefreshData();      // Correct method to refresh the cache
                pivotTable.CalculateData();   // Recalculate the pivot table values

                // Save the workbook with the updated pivot table
                string outputPath = "ShowHiddenColumnFieldDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
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
            ShowHiddenColumnFieldDemo.Run();
        }
    }
}

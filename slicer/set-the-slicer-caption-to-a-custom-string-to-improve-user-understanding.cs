using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCaptionDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate worksheet with sample data for a pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["A3"].PutValue("Vegetable");
                worksheet.Cells["A4"].PutValue("Fruit");
                worksheet.Cells["A5"].PutValue("Vegetable");

                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["B4"].PutValue(150);
                worksheet.Cells["B5"].PutValue(90);

                // Add a table to the data range (required for slicer creation)
                int tableIndex = worksheet.ListObjects.Add("Table1", "A1:B5", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Create a pivot table based on the table data.
                // Use the data range address directly because ListObject.Name may not be available in all versions.
                int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D2", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.CalculateData();

                // Add a slicer linked to the pivot table for the "Category" field.
                // Parameters: pivot table, column index, row index, height (in rows).
                int slicerIndex = worksheet.Slicers.Add(pivotTable, 0, 1, 4);
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Set a custom caption for the slicer
                slicer.Caption = "Select Category";
                slicer.ShowCaption = true; // Ensure the caption header is visible

                // Save the workbook to a file
                string outputPath = "SlicerWithCustomCaption.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
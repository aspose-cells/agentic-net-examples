using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCaptionDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate worksheet with sample data for a pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Amount");
                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["A3"].PutValue("Vegetable");
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["A4"].PutValue("Fruit");
                worksheet.Cells["B4"].PutValue(150);
                worksheet.Cells["A5"].PutValue("Vegetable");
                worksheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D2", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivotTable.CalculateData();

                // Add a slicer linked to the pivot table for the "Category" field
                // Note: The Add method expects destination cell first, then the field name
                int slicerIndex = worksheet.Slicers.Add(pivotTable, "E2", "Category");
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Set a custom caption for the slicer
                slicer.Caption = "Select Category";

                // Ensure the caption header is visible (default is true)
                slicer.ShowCaption = true;

                // Save the workbook to a file
                workbook.Save("SlicerWithCustomCaption.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
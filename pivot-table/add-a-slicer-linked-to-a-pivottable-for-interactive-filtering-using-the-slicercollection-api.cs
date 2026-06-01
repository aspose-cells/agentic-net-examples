using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerDemo
{
    public class SlicerLinkedToPivotDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Quantity";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 10;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 5;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 8;

            // Add a pivot table using the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Add a slicer linked to the pivot table for the "Fruit" field
            SlicerCollection slicers = sheet.Slicers;
            int slicerIndex = slicers.Add(pivot, "F1", "Fruit");
            Slicer slicer = slicers[slicerIndex];

            // Optional: customize slicer appearance
            slicer.Caption = "Fruit Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // Save the workbook with the slicer attached
            string outputPath = "SlicerLinkedToPivot.xlsx";
            workbook.Save(outputPath);
        }
    }
}
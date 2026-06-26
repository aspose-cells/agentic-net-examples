using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;   // For Shape class

class SlicerToImageDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["B1"].Value = "Year";
            sheet.Cells["C1"].Value = "Amount";

            string[] fruits = { "Apple", "Banana", "Apple", "Banana", "Apple", "Banana" };
            int[] years = { 2020, 2020, 2021, 2021, 2022, 2022 };
            int[] amounts = { 50, 70, 60, 80, 55, 75 };

            for (int i = 0; i < fruits.Length; i++)
            {
                sheet.Cells[i + 1, 0].Value = fruits[i];
                sheet.Cells[i + 1, 1].Value = years[i];
                sheet.Cells[i + 1, 2].Value = amounts[i];
            }

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:C7", "E3", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "Fruit", "F12");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Obtain the slicer's shape to determine its bounds
            Shape slicerShape = slicer.Shape;
            int startRow = slicerShape.UpperLeftRow;
            int startCol = slicerShape.UpperLeftColumn;
            int endRow = slicerShape.LowerRightRow;
            int endCol = slicerShape.LowerRightColumn;

            // Convert the bounds to an address string (e.g., "F12:G15")
            string startCell = CellsHelper.CellIndexToName(startRow, startCol);
            string endCell = CellsHelper.CellIndexToName(endRow, endCol);
            string slicerRange = $"{startCell}:{endCell}";

            // Set the worksheet's print area to the slicer bounds
            sheet.PageSetup.PrintArea = slicerRange;

            // Configure image rendering options to output only the defined area
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnlyArea = true // Render only the print area
            };

            // Render the worksheet (which now has the slicer as its print area) to an image file
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "SlicerImage.png");

            // Optional: Save the workbook for verification
            workbook.Save("SlicerDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;   // Pivot table related types

namespace SlicerRenderDemo
{
    class Program
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

                // Add a pivot table based on the data range
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("A1:C7", "E3", "FruitPivot");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table for the "Fruit" field
                SlicerCollection slicers = sheet.Slicers;
                int slicerIdx = slicers.Add(pivot, "Fruit", "F2");
                Slicer slicer = slicers[slicerIdx];

                // Ensure the slicer is printable (use Shape.IsPrintable as IsPrintable is obsolete)
                slicer.Shape.IsPrintable = true;

                // Retrieve the slicer's shape to determine its bounds
                var slicerShape = slicer.Shape;

                int ulRow = slicerShape.UpperLeftRow;      // Upper‑left row index (0‑based)
                int ulCol = slicerShape.UpperLeftColumn;   // Upper‑left column index (0‑based)
                int lrRow = slicerShape.LowerRightRow;     // Lower‑right row index (0‑based)
                int lrCol = slicerShape.LowerRightColumn;  // Lower‑right column index (0‑based)

                // Convert the corner cells to their A1 style names
                string startCell = sheet.Cells[ulRow, ulCol].Name;
                string endCell   = sheet.Cells[lrRow, lrCol].Name;

                // Set the worksheet's print area to exactly the slicer's bounds
                sheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

                // Configure image rendering options: render only the defined area
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true,
                    OnlyArea = true               // Render only the print area without scaling
                };

                // Render the worksheet (which now contains only the slicer) to an image file
                SheetRender renderer = new SheetRender(sheet, options);
                string imagePath = "SlicerImage.png";
                renderer.ToImage(0, imagePath);

                // Save the workbook (optional, to verify the slicer and print area)
                string workbookPath = "SlicerWorkbook.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Slicer rendered to image file '{imagePath}'.");
                Console.WriteLine($"Workbook saved as '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

namespace ExportSlicersToPdf
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create workbook and worksheet --------------------
                Workbook workbook = new Workbook();                     // create a new workbook
                Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet

                // -------------------- Populate sample data --------------------
                sheet.Cells["A1"].PutValue("Fruit");
                sheet.Cells["B1"].PutValue("Year");
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(2020);
                sheet.Cells["C2"].PutValue(120);

                sheet.Cells["A3"].PutValue("Apple");
                sheet.Cells["B3"].PutValue(2021);
                sheet.Cells["C3"].PutValue(150);

                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(2020);
                sheet.Cells["C4"].PutValue(80);

                sheet.Cells["A5"].PutValue("Banana");
                sheet.Cells["B5"].PutValue(2021);
                sheet.Cells["C5"].PutValue(95);

                // -------------------- Create a pivot table --------------------
                // The pivot will be placed starting at cell E2
                int pivotIdx = sheet.PivotTables.Add("A1:C5", "E2", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                // Add fields to the pivot
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------- Add a slicer linked to the pivot --------------------
                // Place the slicer with its upper‑left corner at cell G2
                int slicerIdx = sheet.Slicers.Add(pivot, "G2", "Fruit");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Ensure the slicer is printable (it will be rendered as a static image)
                slicer.Shape.IsPrintable = true;   // use the underlying Shape object

                // -------------------- Configure PDF save options --------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Optional: export only the first sheet (index 0)
                    SheetSet = new SheetSet(new int[] { 0 })
                };

                // -------------------- Save the workbook as PDF --------------------
                string outputPath = "Workbook_With_Slicers.pdf";

                // Ensure we can write to the target location
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerToPdfReport
{
    class Program
    {
        static void Main()
        {
            // 1. Create a workbook and add sample data
            Workbook sourceWb = new Workbook();
            Worksheet sourceWs = sourceWb.Worksheets[0];
            sourceWs.Cells["A1"].PutValue("Category");
            sourceWs.Cells["A2"].PutValue("Fruit");
            sourceWs.Cells["A3"].PutValue("Vegetable");
            sourceWs.Cells["B1"].PutValue("Amount");
            sourceWs.Cells["B2"].PutValue(120);
            sourceWs.Cells["B3"].PutValue(80);

            // 2. Create a pivot table based on the data
            int pivotIdx = sourceWs.PivotTables.Add("A1:B3", "D1", "PivotTable1");
            PivotTable pivot = sourceWs.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field

            // 3. Add a slicer linked to the pivot table (field index 0 = Category)
            int slicerIdx = sourceWs.Slicers.Add(pivot, 20, 2, 0);
            Slicer slicer = sourceWs.Slicers[slicerIdx];
            slicer.IsPrintable = true; // ensure it appears in rendered image

            // 4. Render the worksheet (including the slicer) to an image stream
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };
            SheetRender sheetRender = new SheetRender(sourceWs, imgOptions);
            using (MemoryStream slicerImageStream = new MemoryStream())
            {
                // Render first (and only) page to the stream
                sheetRender.ToImage(0, slicerImageStream);
                slicerImageStream.Position = 0; // reset for reading

                // 5. Create a new workbook that will serve as the PDF report
                Workbook reportWb = new Workbook();
                Worksheet reportWs = reportWb.Worksheets[0];
                reportWs.Name = "Report";

                // 6. Insert the slicer image into the report worksheet
                // Place the image at cell A1 (row 0, column 0)
                reportWs.Pictures.Add(0, 0, slicerImageStream);

                // 7. Save the report workbook as PDF
                reportWb.Save("SlicerReport.pdf", SaveFormat.Pdf);
            }

            // Clean up renderers
            sheetRender.Dispose();
        }
    }
}
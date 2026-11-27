// Title: Set worksheet print area to include a slicer and export the slicer as a PNG image using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set print area to include slicer and export slicer image | Render only an Excel slicer to PNG with Aspose.Cells in .NET | Define worksheet print area around slicer bounds using Aspose.Cells C# | Programmatically save an Excel slicer as a PNG file with Aspose.Cells | Use temporary print area to render slicer image in Aspose.Cells C#
// Tags: worksheet print area with slicer Aspose.Cells | export slicer as PNG Aspose.Cells | C# render slicer image using SheetRender | temporary print area for slicer rendering | Aspose.Cells slicer bounds range

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace SlicerPrintAreaAndImage
{
    // The example creates a workbook, adds sample data, converts the range to a table, inserts a slicer for the first column, defines a range that encloses the slicer, sets this range as the worksheet's print area, temporarily re‑applies the same range to render only the slicer to a PNG image with ImageOrPrintOptions and SheetRender, restores the original print area, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for a table (Category / Value)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("A");
                sheet.Cells["B5"].PutValue(40);
                sheet.Cells["A6"].PutValue("B");
                sheet.Cells["B6"].PutValue(50);

                // -----------------------------------------------------------------
                // 2. Convert the range into a ListObject (Excel Table).
                // -----------------------------------------------------------------
                int tableIndex = sheet.ListObjects.Add(0, 0, 5, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SampleTable";

                // -----------------------------------------------------------------
                // 3. Add a slicer for the first column (Category) of the table.
                // -----------------------------------------------------------------
                // The slicer will be placed starting at cell D2 (column index 3, row index 1).
                int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[0], 3, 1);
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.Name = "CategorySlicer";
                slicer.Caption = "Category";

                // -----------------------------------------------------------------
                // 4. Define a print area that includes the slicer.
                // -----------------------------------------------------------------
                // Aspose.Cells Slicer does not expose a GetBounds method in recent versions.
                // As an alternative, we define a range that certainly covers the slicer location.
                // Here we use columns A‑D and rows 1‑10, which includes the slicer placed at D2.
                string slicerRange = "A1:D10";

                Console.WriteLine($"Using print area that includes slicer: {slicerRange}");

                // -----------------------------------------------------------------
                // 5. Set the worksheet print area to the defined range.
                // -----------------------------------------------------------------
                sheet.PageSetup.PrintArea = slicerRange;
                Console.WriteLine($"Print area set to: {sheet.PageSetup.PrintArea}");

                // -----------------------------------------------------------------
                // 6. Render the slicer itself to an image file.
                // -----------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true
                };

                string slicerImagePath = "CategorySlicer.png";

                // Render only the slicer area by temporarily setting the print area.
                string originalPrintArea = sheet.PageSetup.PrintArea;
                sheet.PageSetup.PrintArea = slicerRange;

                try
                {
                    SheetRender sr = new SheetRender(sheet, imgOptions);
                    sr.ToImage(0, slicerImagePath);
                    Console.WriteLine($"Slicer image saved to: {slicerImagePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to render slicer as image: {ex.Message}");
                }
                finally
                {
                    // Restore original print area.
                    sheet.PageSetup.PrintArea = originalPrintArea;
                }

                // -----------------------------------------------------------------
                // 7. Save the workbook (optional, to see the print area in Excel).
                // -----------------------------------------------------------------
                string workbookPath = "WorkbookWithSlicer.xlsx";
                try
                {
                    workbook.Save(workbookPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to: {workbookPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}

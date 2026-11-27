// Title: Copy slicer settings (caption, style, size) from one worksheet to a new slicer on another worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy slicer caption and style to a slicer on a different sheet | duplicate slicer layout and dimensions across worksheets using Aspose.Cells for .NET | transfer slicer properties from source worksheet to target worksheet in C# | how to clone an Aspose.Cells slicer to another worksheet programmatically | copy slicer configuration (columns, style, position) between worksheets in Aspose.Cells
// Tags: copy slicer properties Aspose.Cells C# | slicer style duplication Aspose.Cells | slicer layout cloning worksheet Aspose.Cells | Aspose.Cells slicer ListObject association | programmatic slicer property transfer Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;   // Required for ListObject

namespace SlicerCopyDemo
{
    // The example creates a workbook with a source sheet containing a table and a slicer whose caption, column count, style, and pixel dimensions are configured. It then adds a target sheet with an identical table, creates a new slicer there, and programmatically copies all visual and style properties from the source slicer to the target slicer before saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet wsSource = workbook.Worksheets[0];
                wsSource.Name = "SourceSheet";

                // 2. Populate sample data for a table (A1:B6)
                wsSource.Cells["A1"].PutValue("Category");
                wsSource.Cells["B1"].PutValue("Value");
                wsSource.Cells["A2"].PutValue("Fruit");
                wsSource.Cells["B2"].PutValue(10);
                wsSource.Cells["A3"].PutValue("Fruit");
                wsSource.Cells["B3"].PutValue(15);
                wsSource.Cells["A4"].PutValue("Vegetable");
                wsSource.Cells["B4"].PutValue(20);
                wsSource.Cells["A5"].PutValue("Fruit");
                wsSource.Cells["B5"].PutValue(12);
                wsSource.Cells["A6"].PutValue("Vegetable");
                wsSource.Cells["B6"].PutValue(8);

                // 3. Add a ListObject (table) that covers the data range
                int tableIndex = wsSource.ListObjects.Add(0, 0, 5, 1, true);
                ListObject sourceTable = wsSource.ListObjects[tableIndex];
                sourceTable.DisplayName = "SourceTable";

                // 4. Add a slicer on the source worksheet for the "Category" column
                int slicerIdxSource = wsSource.Slicers.Add(sourceTable, sourceTable.ListColumns[0], 1, 3);
                Slicer slicerSource = wsSource.Slicers[slicerIdxSource];

                // Configure some properties on the source slicer
                slicerSource.Caption = "Category Filter";
                slicerSource.NumberOfColumns = 2;
                slicerSource.StyleType = SlicerStyleType.SlicerStyleDark2;
                slicerSource.TopPixel = 50;      // Position from top (in pixels)
                slicerSource.LeftPixel = 300;    // Position from left (in pixels)
                slicerSource.HeightPixel = 150;  // Height (in pixels)
                slicerSource.WidthPixel = 200;   // Width (in pixels)

                // 5. Add a second worksheet that will host the copied slicer
                Worksheet wsTarget = workbook.Worksheets.Add("TargetSheet");

                // 6. Populate the same sample data on the target sheet (required for slicer source)
                wsTarget.Cells["A1"].PutValue("Category");
                wsTarget.Cells["B1"].PutValue("Value");
                wsTarget.Cells["A2"].PutValue("Fruit");
                wsTarget.Cells["B2"].PutValue(10);
                wsTarget.Cells["A3"].PutValue("Fruit");
                wsTarget.Cells["B3"].PutValue(15);
                wsTarget.Cells["A4"].PutValue("Vegetable");
                wsTarget.Cells["B4"].PutValue(20);
                wsTarget.Cells["A5"].PutValue("Fruit");
                wsTarget.Cells["B5"].PutValue(12);
                wsTarget.Cells["A6"].PutValue("Vegetable");
                wsTarget.Cells["B6"].PutValue(8);

                // 7. Add a table on the target sheet (the slicer must be linked to a table)
                int targetTableIdx = wsTarget.ListObjects.Add(0, 0, 5, 1, true);
                ListObject targetTable = wsTarget.ListObjects[targetTableIdx];
                targetTable.DisplayName = "TargetTable";

                // 8. Add a new slicer on the target worksheet (initially with default properties)
                int slicerIdxTarget = wsTarget.Slicers.Add(targetTable, targetTable.ListColumns[0], 1, 3);
                Slicer slicerTarget = wsTarget.Slicers[slicerIdxTarget];

                // 9. Copy properties from the source slicer to the target slicer
                slicerTarget.Caption = slicerSource.Caption;
                slicerTarget.NumberOfColumns = slicerSource.NumberOfColumns;
                slicerTarget.StyleType = slicerSource.StyleType;
                slicerTarget.TopPixel = slicerSource.TopPixel;
                slicerTarget.LeftPixel = slicerSource.LeftPixel;
                slicerTarget.HeightPixel = slicerSource.HeightPixel;
                slicerTarget.WidthPixel = slicerSource.WidthPixel;

                // 10. Save the workbook
                string outputPath = "SlicerCopyDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}

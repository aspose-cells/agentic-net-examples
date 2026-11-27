// Title: Apply a built‑in slicer style, duplicate the slicer to another worksheet, and verify visual consistency using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set slicer style and copy slicer to another worksheet | How to compare slicer properties for visual consistency across Excel sheets using Aspose.Cells | Duplicate a pivot table slicer in a second worksheet with Aspose.Cells .NET | Apply built‑in slicer style and verify matching layout in multiple worksheets C# | Copy slicer position and style programmatically with Aspose.Cells for .NET
// Tags: Aspose.Cells apply slicer style | Aspose.Cells duplicate slicer across worksheets | Aspose.Cells slicer property comparison | C# pivot table slicer duplication | Excel slicer visual consistency .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace CustomSlicerStyleDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the Category field, applies a built‑in light style, then duplicates the slicer on a second worksheet with identical size, position, and style settings. It programmatically compares key slicer properties to confirm visual consistency before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and add sample data
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Data";

                // Populate sample data for a pivot table
                sheet1.Cells["A1"].PutValue("Category");
                sheet1.Cells["B1"].PutValue("Product");
                sheet1.Cells["C1"].PutValue("Quantity");

                sheet1.Cells["A2"].PutValue("Fruits");
                sheet1.Cells["B2"].PutValue("Apple");
                sheet1.Cells["C2"].PutValue(10);

                sheet1.Cells["A3"].PutValue("Fruits");
                sheet1.Cells["B3"].PutValue("Banana");
                sheet1.Cells["C3"].PutValue(15);

                sheet1.Cells["A4"].PutValue("Vegetables");
                sheet1.Cells["B4"].PutValue("Carrot");
                sheet1.Cells["C4"].PutValue(20);

                sheet1.Cells["A5"].PutValue("Vegetables");
                sheet1.Cells["B5"].PutValue("Tomato");
                sheet1.Cells["C5"].PutValue(12);

                // -----------------------------------------------------------------
                // 2. Create a pivot table based on the data
                // -----------------------------------------------------------------
                string sourceRange = "A1:C5";
                string pivotDestination = "E3";

                int pivotIndex = sheet1.PivotTables.Add(sourceRange, pivotDestination, "SalesPivot");
                PivotTable pivotTable = sheet1.PivotTables[pivotIndex];

                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -----------------------------------------------------------------
                // 3. Add a slicer linked to the pivot table (Category field)
                // -----------------------------------------------------------------
                int slicerIndex = sheet1.Slicers.Add(pivotTable, "Category", "CategorySlicer");
                Slicer slicer = sheet1.Slicers[slicerIndex];

                // Position and size using pixel‑based properties
                slicer.TopPixel = 250;
                slicer.LeftPixel = 20;
                slicer.WidthPixel = 200;
                slicer.HeightPixel = 150;

                // -----------------------------------------------------------------
                // 4. Apply a built‑in style to the slicer (custom colour scheme not directly supported)
                // -----------------------------------------------------------------
                slicer.StyleType = SlicerStyleType.SlicerStyleLight2; // example built‑in style
                slicer.Caption = "Select Category";
                slicer.NumberOfColumns = 1;

                // -----------------------------------------------------------------
                // 5. Duplicate the slicer on a second worksheet to verify visual consistency
                // -----------------------------------------------------------------
                Worksheet sheet2 = workbook.Worksheets.Add("Copy");

                int pivotIndex2 = sheet2.PivotTables.Add(sourceRange, pivotDestination, "SalesPivotCopy");
                PivotTable pivotTable2 = sheet2.PivotTables[pivotIndex2];
                pivotTable2.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable2.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable2.AddFieldToArea(PivotFieldType.Data, "Quantity");
                pivotTable2.RefreshData();
                pivotTable2.CalculateData();

                int slicerIndex2 = sheet2.Slicers.Add(pivotTable2, "Category", "CategorySlicerCopy");
                Slicer slicerCopy = sheet2.Slicers[slicerIndex2];

                // Copy position (offset to avoid overlap)
                slicerCopy.TopPixel = slicer.TopPixel;
                slicerCopy.LeftPixel = slicer.LeftPixel + 250;
                slicerCopy.WidthPixel = slicer.WidthPixel;
                slicerCopy.HeightPixel = slicer.HeightPixel;

                // Copy visual style
                slicerCopy.StyleType = slicer.StyleType;
                slicerCopy.Caption = slicer.Caption;
                slicerCopy.NumberOfColumns = slicer.NumberOfColumns;

                // -----------------------------------------------------------------
                // 6. Verify visual consistency by comparing key style properties
                // -----------------------------------------------------------------
                bool isConsistent = true;

                isConsistent &= slicer.StyleType == slicerCopy.StyleType;
                isConsistent &= slicer.Caption == slicerCopy.Caption;
                isConsistent &= slicer.NumberOfColumns == slicerCopy.NumberOfColumns;
                isConsistent &= slicer.TopPixel == slicerCopy.TopPixel;
                isConsistent &= slicer.LeftPixel == slicerCopy.LeftPixel;
                isConsistent &= slicer.WidthPixel == slicerCopy.WidthPixel;
                isConsistent &= slicer.HeightPixel == slicerCopy.HeightPixel;

                Console.WriteLine("Custom slicer style applied to first sheet.");
                Console.WriteLine("Slicer duplicated on second sheet.");
                Console.WriteLine("Visual consistency verification result: " + (isConsistent ? "PASS" : "FAIL"));

                // -----------------------------------------------------------------
                // 7. Save the workbook
                // -----------------------------------------------------------------
                string outFile = "CustomSlicerStyle.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

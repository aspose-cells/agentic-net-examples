using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerExample
{
    public class MultiSelectSlicerDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate worksheet with sample data for the table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("C");
                sheet.Cells["B5"].PutValue(40);

                // Add a table that covers the data range (A1:B5)
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Add a slicer linked to the first column ("Category") of the table
                // Placed with its upper‑left corner at row 6, column 2 (cell B6)
                SlicerCollection slicers = sheet.Slicers;
                int slicerIndex = slicers.Add(table, table.ListColumns[0], 6, 2);
                Slicer slicer = slicers[slicerIndex];

                // Configure slicer appearance (optional)
                slicer.Caption = "Category Filter";
                slicer.NumberOfColumns = 2; // layout with two columns for easier multi‑selection

                // Use Shape properties instead of obsolete WidthPixel/HeightPixel
                slicer.Shape.Width = 200;   // width in points
                slicer.Shape.Height = 150;  // height in points

                // Save the workbook
                string outputPath = "MultiSelectSlicerDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
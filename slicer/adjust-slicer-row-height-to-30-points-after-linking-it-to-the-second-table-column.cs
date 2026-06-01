using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerExample
{
    public class AdjustSlicerRowHeight
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue("Fruit");
                sheet.Cells["A3"].PutValue("Carrot");
                sheet.Cells["B3"].PutValue("Vegetable");
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue("Fruit");

                // Add a ListObject (table) covering the data range
                int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Add a slicer linked to the second column (Category) of the table
                int slicerRow = 6;      // zero‑based index (row 7)
                int slicerColumn = 0;   // column A
                int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[1], slicerRow, slicerColumn);
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Adjust the slicer row height to 30 points
                slicer.RowHeight = 30;

                // Save the workbook
                string outputPath = "SlicerRowHeightAdjusted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustSlicerRowHeight.Run();
        }
    }
}
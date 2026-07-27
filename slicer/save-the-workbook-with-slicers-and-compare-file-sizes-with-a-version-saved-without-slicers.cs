using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerSizeComparison
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for a table
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Year");
                cells["C1"].PutValue("Amount");

                cells["A2"].PutValue("A");
                cells["B2"].PutValue(2020);
                cells["C2"].PutValue(10);

                cells["A3"].PutValue("B");
                cells["B3"].PutValue(2020);
                cells["C3"].PutValue(20);

                cells["A4"].PutValue("A");
                cells["B4"].PutValue(2021);
                cells["C4"].PutValue(30);

                cells["A5"].PutValue("B");
                cells["B5"].PutValue(2021);
                cells["C5"].PutValue(40);

                // Add a ListObject (table) using the data range
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Set a display name for the table (Name property may not be available in some versions)
                table.DisplayName = "DataTable";

                // Add a slicer for the first column of the table
                int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[0], "E1");
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.Caption = "Category Slicer";

                // Save workbook with slicer
                string withSlicerPath = "Workbook_WithSlicers.xlsx";
                workbook.Save(withSlicerPath, SaveFormat.Xlsx);

                // Remove all slicers from the worksheet
                sheet.Slicers.Clear();

                // Save workbook without slicer
                string withoutSlicerPath = "Workbook_WithoutSlicers.xlsx";
                workbook.Save(withoutSlicerPath, SaveFormat.Xlsx);

                // Compare file sizes (ensure files exist before accessing)
                long sizeWithSlicer = File.Exists(withSlicerPath) ? new FileInfo(withSlicerPath).Length : 0;
                long sizeWithoutSlicer = File.Exists(withoutSlicerPath) ? new FileInfo(withoutSlicerPath).Length : 0;

                Console.WriteLine($"File size with slicers: {sizeWithSlicer} bytes");
                Console.WriteLine($"File size without slicers: {sizeWithoutSlicer} bytes");
                Console.WriteLine($"Difference: {sizeWithSlicer - sizeWithoutSlicer} bytes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
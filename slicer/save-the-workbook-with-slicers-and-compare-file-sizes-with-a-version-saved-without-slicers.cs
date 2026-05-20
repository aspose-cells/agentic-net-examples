using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerSizeComparison
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
                Cells cells = sheet.Cells;

                // Populate sample data for a table (A1:C9)
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Year");
                cells["C1"].PutValue("Amount");

                string[] categories = { "A", "B", "A", "B", "A", "B", "A", "B", "A" };
                int[] years = { 2020, 2020, 2020, 2020, 2020, 2021, 2021, 2021, 2021 };
                int[] amounts = { 10, 20, 30, 40, 50, 60, 70, 80, 90 };

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[i + 1, 0].PutValue(categories[i]); // Column A
                    cells[i + 1, 1].PutValue(years[i]);     // Column B
                    cells[i + 1, 2].PutValue(amounts[i]);  // Column C
                }

                // Add a table (ListObject) covering the data range
                int tableIndex = sheet.ListObjects.Add(0, 0, categories.Length, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Use DisplayName instead of Name for compatibility with older Aspose.Cells versions
                table.DisplayName = "DataTable";

                // Add a slicer for the first column (Category) of the table, positioned at cell E2
                int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[0], "E2");
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.Caption = "Category Slicer";

                // Save workbook with slicer
                string withSlicerPath = "WorkbookWithSlicer.xlsx";
                workbook.Save(withSlicerPath, SaveFormat.Xlsx);
                long sizeWithSlicer = new FileInfo(withSlicerPath).Length;

                // Remove all slicers from the worksheet
                sheet.Slicers.Clear();

                // Save workbook without slicer
                string withoutSlicerPath = "WorkbookWithoutSlicer.xlsx";
                workbook.Save(withoutSlicerPath, SaveFormat.Xlsx);
                long sizeWithoutSlicer = new FileInfo(withoutSlicerPath).Length;

                // Output the comparison results
                Console.WriteLine($"File size with slicer    : {sizeWithSlicer} bytes");
                Console.WriteLine($"File size without slicer : {sizeWithoutSlicer} bytes");
                Console.WriteLine($"Size difference          : {sizeWithSlicer - sizeWithoutSlicer} bytes");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
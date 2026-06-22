using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
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

                // Populate sample data (headers + rows)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Year");
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(2020);
                sheet.Cells["C2"].PutValue(50);

                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue(2021);
                sheet.Cells["C3"].PutValue(70);

                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue(2020);
                sheet.Cells["C4"].PutValue(30);

                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue(2021);
                sheet.Cells["C5"].PutValue(45);

                // Convert the data range into a ListObject (Excel table)
                // Parameters: firstRow, firstColumn, totalRows-1, totalColumns-1, hasHeaders
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Use DisplayName instead of Name for compatibility with older Aspose.Cells versions
                table.DisplayName = "DataTable";

                // Apply an AutoFilter to the table and filter "Category" = "Fruit"
                sheet.AutoFilter.SetRange(0, 0, 2); // Apply filter to columns A‑C, starting at row 0 (header row)
                sheet.AutoFilter.AddFilter(0, "Fruit"); // Column 0 (Category) filter
                sheet.AutoFilter.Refresh();

                // Create a named range that refers to the filtered data range of the table
                // Note: The DataRange includes the header row; you can adjust if needed.
                Aspose.Cells.Range dataRange = table.DataRange;
                string address = $"{sheet.Name}!{CellsHelper.CellIndexToName(dataRange.FirstRow, dataRange.FirstColumn)}:" +
                                 $"{CellsHelper.CellIndexToName(dataRange.FirstRow + dataRange.RowCount - 1, dataRange.FirstColumn + dataRange.ColumnCount - 1)}";

                int nameIdx = workbook.Worksheets.Names.Add("FilteredData");
                workbook.Worksheets.Names[nameIdx].RefersTo = "=" + address;

                // Add a slicer linked to the first column ("Category") of the ListObject
                // Destination cell for the slicer's upper‑left corner is E2
                int slicerIdx = sheet.Slicers.Add(table, 0, "E2");
                Slicer slicer = sheet.Slicers[slicerIdx];
                slicer.Caption = "Category Slicer";
                slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

                // Save the workbook (ensure the directory exists)
                string outputPath = "NamedRangeSlicerDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
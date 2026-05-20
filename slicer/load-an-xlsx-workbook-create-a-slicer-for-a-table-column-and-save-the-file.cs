using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one table (ListObject) in the worksheet
            if (sheet.ListObjects.Count == 0)
            {
                // Create sample data for a table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(30);

                // Add a table covering the range A1:B4 (firstRow, firstColumn, totalRows-1, totalColumns-1)
                int tableIdx = sheet.ListObjects.Add(0, 0, 3, 1, true);
                // No explicit Refresh method needed; the table is ready after creation
            }

            // Retrieve the first table in the worksheet
            ListObject table = sheet.ListObjects[0];

            // Add a slicer for the first column of the table and place it at cell E2
            int slicerIdx = sheet.Slicers.Add(table, 0, "E2");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.Caption = "Category Slicer";

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
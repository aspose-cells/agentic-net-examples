// Title: Add a slicer linked to a ListObject column and position it at the worksheet’s top‑right corner using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add slicer to a table column and set its location | how to place a slicer in the top right corner of an Excel sheet with Aspose.Cells | programmatically associate a slicer with a ListObject column using Aspose.Cells for .NET | C# example for creating a slicer and moving it two columns right of existing data | Aspose.Cells slicer placement top right worksheet code sample
// Tags: Aspose.Cells ListObject slicer creation | C# set slicer position top right worksheet | Aspose.Cells associate slicer with table column | programmatic slicer placement Excel C# | Aspose.Cells slicer style configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace SlicerDemo
{
    // The example creates a new workbook, fills a range with sample data, converts it into a ListObject (table), adds a slicer linked to the first table column (Category), positions the slicer two columns to the right of the data on the top row (top‑right corner), optionally applies a slicer style, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table (A1:C6)
                Cells cells = worksheet.Cells;
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Year");
                cells["C1"].PutValue("Amount");

                string[] categories = { "Apple", "Banana", "Cherry", "Apple", "Banana", "Cherry" };
                int[] years = { 2020, 2020, 2020, 2021, 2021, 2021 };
                int[] amounts = { 50, 70, 90, 60, 80, 100 };

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[i + 1, 0].PutValue(categories[i]);   // Column A
                    cells[i + 1, 1].PutValue(years[i]);       // Column B
                    cells[i + 1, 2].PutValue(amounts[i]);     // Column C
                }

                // Define the range that will become a table (A1:C7)
                int firstRow = 0;
                int firstColumn = 0;
                int totalRows = categories.Length + 1; // + header
                int totalColumns = 3;

                // Add a ListObject (table) to the worksheet
                int tableIndex = worksheet.ListObjects.Add(firstRow, firstColumn, totalRows - 1, totalColumns - 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.ShowTableStyleFirstColumn = true;
                table.ShowTableStyleLastColumn = true;
                table.TableStyleType = TableStyleType.TableStyleMedium9;

                // Determine a column index that is far enough to the right to act as "top‑right"
                int lastUsedColumn = worksheet.Cells.MaxColumn; // zero‑based index
                int slicerColumn = lastUsedColumn + 2; // two columns to the right of existing data
                int slicerRow = 0; // top row

                // Add a slicer linked to the first column ("Category") of the table
                int slicerIndex = worksheet.Slicers.Add(table, table.ListColumns[0], slicerRow, slicerColumn);
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Optionally set a style for better visibility (if supported by the version)
                // slicer.SlicerStyle = SlicerStyleType.SlicerStyleLight1; // Uncomment if property exists

                // Save the workbook
                string outputPath = "SlicerTopRightDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the slicer:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

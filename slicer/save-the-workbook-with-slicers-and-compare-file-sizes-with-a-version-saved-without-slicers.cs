// Title: Measure the XLSX file size difference when saving an Aspose.Cells workbook with and without a slicer in C#
// AI Prompts: Generate C# code that creates a workbook, adds a ListObject table, inserts a slicer for a column, saves the file, clears all slicers, saves a second file, and prints the byte sizes of both files. | Write a reusable C# method that takes a Worksheet and column index, adds a slicer, saves the workbook twice (with and without the slicer), and returns the size delta in bytes. | Adapt the example to use a pivot table slicer, save the workbook, and output the size difference in kilobytes.
// Common Searches: how much does a slicer increase the size of an .xlsx file using Aspose.Cells | C# Aspose.Cells example to compare workbook size with slicer versus without | remove all slicers from a worksheet programmatically Aspose.Cells .NET | measure file size impact of adding a slicer to an Excel workbook in C#
// Tags: Aspose.Cells add slicer to ListObject | C# save workbook with slicer XLSX | Aspose.Cells clear slicers worksheet | XLSX file size impact slicer Aspose | compare workbook size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerSizeComparison
{
    // // Demonstrates creating a workbook with sample data, adding a table and a slicer, saving the file, clearing all slicers, saving a second file, and printing the byte sizes and their difference.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("A");
            cells["B4"].PutValue(30);
            cells["A5"].PutValue("B");
            cells["B5"].PutValue(40);

            // Add a table covering the data range
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Add a slicer for the first column of the table
            // The slicer will be placed at cell D1
            int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[0], "D1");
            Slicer slicer = sheet.Slicers[slicerIndex];
            slicer.Caption = "Category Slicer";

            // Save workbook with slicer
            string fileWithSlicer = "WorkbookWithSlicer.xlsx";
            workbook.Save(fileWithSlicer, SaveFormat.Xlsx);
            long sizeWithSlicer = new FileInfo(fileWithSlicer).Length;

            // Remove all slicers from the worksheet
            sheet.Slicers.Clear();

            // Save workbook without slicer
            string fileWithoutSlicer = "WorkbookWithoutSlicer.xlsx";
            workbook.Save(fileWithoutSlicer, SaveFormat.Xlsx);
            long sizeWithoutSlicer = new FileInfo(fileWithoutSlicer).Length;

            // Output file sizes and difference
            Console.WriteLine($"Size with slicer    : {sizeWithSlicer} bytes");
            Console.WriteLine($"Size without slicer : {sizeWithoutSlicer} bytes");
            Console.WriteLine($"Difference          : {sizeWithSlicer - sizeWithoutSlicer} bytes");
        }
    }
}

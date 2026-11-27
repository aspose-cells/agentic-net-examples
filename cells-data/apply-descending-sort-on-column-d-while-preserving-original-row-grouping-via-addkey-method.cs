// Title: Sort Excel data by column D descending while keeping groups together using Aspose.Cells DataSorter.AddKey in C#
// Common Searches: C# Aspose.Cells sort rows by column D descending within each group | How to preserve group order when sorting an Excel sheet with Aspose.Cells DataSorter | Using AddKey to sort multiple columns in Aspose.Cells .NET example | Aspose.Cells sort by group column then score column descending | Define CellArea for sorting a specific range in Aspose.Cells C#
// Tags: DataSorter multi-key sorting | C# descending sort using AddKey | Excel column D descending sort | preserve row groups during sort | CellArea range for sorting

using System;
using Aspose.Cells;

namespace AsposeCellsSortExample
{
    // This C# example creates a workbook, adds sample data with a group identifier in column A and a score in column D, then uses Aspose.Cells DataSorter to first sort on the group column (A) ascending and then on the score column (D) descending. The sort range is defined with a CellArea covering the header and data rows, the workbook is saved as SortedByColumnD.xlsx, and the sorted rows are printed to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to output file
            const string outputPath = "SortedByColumnD.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate sample data
            // Column A : Group identifier (to preserve grouping)
            // Column B : Item name
            // Column C : Some numeric value
            // Column D : Value to sort descending within each group
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Group");
            sheet.Cells["B1"].PutValue("Item");
            sheet.Cells["C1"].PutValue("Qty");
            sheet.Cells["D1"].PutValue("Score");

            // Sample rows
            // Group 1
            sheet.Cells["A2"].PutValue("G1");
            sheet.Cells["B2"].PutValue("Item A");
            sheet.Cells["C2"].PutValue(10);
            sheet.Cells["D2"].PutValue(85);

            sheet.Cells["A3"].PutValue("G1");
            sheet.Cells["B3"].PutValue("Item B");
            sheet.Cells["C3"].PutValue(15);
            sheet.Cells["D3"].PutValue(92);

            sheet.Cells["A4"].PutValue("G1");
            sheet.Cells["B4"].PutValue("Item C");
            sheet.Cells["C4"].PutValue(8);
            sheet.Cells["D4"].PutValue(78);

            // Group 2
            sheet.Cells["A5"].PutValue("G2");
            sheet.Cells["B5"].PutValue("Item D");
            sheet.Cells["C5"].PutValue(20);
            sheet.Cells["D5"].PutValue(67);

            sheet.Cells["A6"].PutValue("G2");
            sheet.Cells["B6"].PutValue("Item E");
            sheet.Cells["C6"].PutValue(12);
            sheet.Cells["D6"].PutValue(88);

            sheet.Cells["A7"].PutValue("G2");
            sheet.Cells["B7"].PutValue("Item F");
            sheet.Cells["C7"].PutValue(5);
            sheet.Cells["D7"].PutValue(73);

            // ------------------------------------------------------------
            // Configure the DataSorter
            // ------------------------------------------------------------
            // 1. First key – column A (Group) – ascending to keep groups together
            // 2. Second key – column D (Score) – descending within each group
            // The AddKey method signature:
            //   void AddKey(int columnIndex, SortOnType sortOn, SortOrder order, object customList)
            // columnIndex is zero‑based.
            // ------------------------------------------------------------
            DataSorter sorter = workbook.DataSorter;

            // Preserve original grouping by sorting on Group column (A) ascending
            sorter.AddKey(0, SortOnType.Value, SortOrder.Ascending, null);

            // Then sort on Score column (D) descending
            sorter.AddKey(3, SortOnType.Value, SortOrder.Descending, null);

            // Define the range to be sorted (including header row)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,      // Row 0 = header
                EndRow = 6,        // Last data row (zero‑based)
                StartColumn = 0,   // Column A
                EndColumn = 3      // Column D
            };

            // Perform the sort
            sorter.Sort(sheet.Cells, sortArea);

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            try
            {
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }

            // Optional: display sorted data to console for verification
            Console.WriteLine("\nSorted data (Group | Item | Qty | Score):");
            for (int row = 1; row <= sortArea.EndRow; row++)
            {
                string group = sheet.Cells[row, 0].StringValue;
                string item = sheet.Cells[row, 1].StringValue;
                string qty = sheet.Cells[row, 2].StringValue;
                string score = sheet.Cells[row, 3].StringValue;
                Console.WriteLine($"{group} | {item} | {qty} | {score}");
            }
        }
    }
}

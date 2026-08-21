// Title: Remove duplicate rows in Aspose.Cells (C#) using a composite key of two columns
// Description: Demonstrates how to use Aspose.Cells' Cells.RemoveDuplicates method to delete rows that share the same values in the ID and Name columns, preserving the first occurrence and saving the workbook.
// Keywords: Aspose.Cells | C# RemoveDuplicates | duplicate rows | composite key | Excel deduplication | RemoveDuplicates method | ID and Name | Aspose.Cells example | Excel data cleaning | Aspose.Cells Cells.RemoveDuplicates
// Common Searches: Aspose.Cells remove duplicate rows C# | RemoveDuplicates composite key Aspose.Cells | Delete duplicate Excel rows using two columns Aspose | C# Aspose.Cells deduplicate based on ID and Name | How to use Cells.RemoveDuplicates with multiple columns
// Developer Intent: The developer needs to eliminate rows that have identical values in both the ID and Name columns of an Excel sheet using Aspose.Cells' built‑in RemoveDuplicates functionality.
// Use Cases: Clean imported CSV data before analysis by removing repeated ID‑Name pairs. | Prepare a customer report that lists each unique customer (ID + Name) only once. | Compress large financial spreadsheets by deduplicating rows based on a composite key.
// AI Prompts: Write C# code that removes duplicate rows in an Aspose.Cells worksheet based on three columns (e.g., ID, Name, Date) while keeping the row with the highest Value. | Explain the purpose of the columnOffsets parameter in Cells.RemoveDuplicates and show how to specify non‑adjacent columns. | Provide a sample that removes duplicates from a worksheet containing an Excel table, preserving the table’s style and formulas.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to use Aspose.Cells' Cells.RemoveDuplicates method to delete rows that share the same values in the ID and Name columns, preserving the first occurrence and saving the workbook.
    public class RemoveDuplicateRowsCompositeKeyDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Value");

            // Add sample data with duplicate composite keys (ID + Name)
            // Row 2
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("John");
            cells["C2"].PutValue(100);
            // Row 3 (duplicate of row 2 based on ID and Name)
            cells["A3"].PutValue(1);
            cells["B3"].PutValue("John");
            cells["C3"].PutValue(200);
            // Row 4
            cells["A4"].PutValue(2);
            cells["B4"].PutValue("Jane");
            cells["C4"].PutValue(150);
            // Row 5 (duplicate of row 4 based on ID and Name)
            cells["A5"].PutValue(2);
            cells["B5"].PutValue("Jane");
            cells["C5"].PutValue(250);
            // Row 6 (unique)
            cells["A6"].PutValue(3);
            cells["B6"].PutValue("Bob");
            cells["C6"].PutValue(300);

            // Determine the range that contains data
            int startRow = 1; // data starts after header (0‑based index)
            int startColumn = 0; // column A
            int endRow = cells.MaxDataRow; // last row with data
            int endColumn = 2; // column C (0‑based)

            // Remove duplicates based on the first two columns (ID and Name)
            // hasHeaders = true because row 0 contains column titles
            // columnOffsets = new int[] { 0, 1 } specifies that columns A and B form the composite key
            cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, true, new int[] { 0, 1 });

            // Prepare output file path
            string outputPath = "RemoveDuplicatesCompositeKey.xlsx";

            // Save the result
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}

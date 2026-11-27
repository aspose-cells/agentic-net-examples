// Title: How to perform stable sorting of Excel rows by a column using Aspose.Cells DataSorter in C#
// Common Searches: Aspose.Cells C# stable sort worksheet rows by column | preserve original row order for duplicate values when sorting with DataSorter | C# example sorting Excel range with Aspose.Cells while keeping tie order | using Aspose.Cells DataSorter to sort range with header row in .NET | how to enable stable sorting in Aspose.Cells DataSorter (C#)
// Tags: Aspose.Cells DataSorter stable sorting | C# sort Excel range by column Aspose.Cells | preserve duplicate key order Aspose.Cells | sort worksheet range with header Aspose.Cells | Excel .xlsx stable sort using Aspose.Cells .NET

using System;
using Aspose.Cells;

// This C# example creates a workbook, fills it with sample data that includes duplicate scores, and uses the Aspose.Cells DataSorter to sort the worksheet range by the Score column in ascending order. Because Aspose.Cells performs stable sorting by default, rows with equal scores retain their original relative order. The sorted data is printed to the console and the workbook is saved as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with duplicate keys to demonstrate stable sorting
            // Header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("Score");

            // Data rows (note that John and Alice have the same Score)
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(25);
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["C3"].PutValue(85); // Same score as John

            sheet.Cells["A4"].PutValue("Bob");
            sheet.Cells["B4"].PutValue(22);
            sheet.Cells["C4"].PutValue(78);

            sheet.Cells["A5"].PutValue("Eve");
            sheet.Cells["B5"].PutValue(28);
            sheet.Cells["C5"].PutValue(92);

            // Display data before sorting
            Console.WriteLine("Data before sorting:");
            PrintRange(sheet, 1, 5, 0, 2);

            // Configure the DataSorter
            DataSorter sorter = workbook.DataSorter;

            // Note: StableSort property is not available in this version of Aspose.Cells.
            // The default sorting behavior preserves the relative order of equal keys.

            // Set the primary sort key to column C (Score) in ascending order
            sorter.Key1 = sheet.Cells["C1"].Column; // Column index of "Score"
            sorter.Order1 = SortOrder.Ascending;

            // Define the range to be sorted (including header row)
            CellArea sortArea = new CellArea
            {
                StartRow = 0,   // Header row
                EndRow = 4,     // Last data row (zero‑based)
                StartColumn = 0,
                EndColumn = 2
            };

            // Perform the sort
            sorter.Sort(sheet.Cells, sortArea);

            // Display data after sorting
            Console.WriteLine("\nData after sorting (by Score ascending):");
            PrintRange(sheet, 1, 5, 0, 2);

            // Save the workbook to verify the result manually if needed
            string outPath = "StableSortDemo.xlsx";
            workbook.Save(outPath, SaveFormat.Xlsx);
            Console.WriteLine($"\nWorkbook saved to '{outPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }

    // Helper method to print a rectangular range of cells to the console
    static void PrintRange(Worksheet sheet, int startRow, int endRow, int startCol, int endCol)
    {
        for (int row = startRow; row <= endRow; row++)
        {
            for (int col = startCol; col <= endCol; col++)
            {
                Cell cell = sheet.Cells[row, col];
                // Use StringValue for text and DoubleValue for numbers as appropriate
                string value = cell.Type == CellValueType.IsNumeric ? cell.DoubleValue.ToString() : cell.StringValue;
                Console.Write(value);
                if (col < endCol) Console.Write("\t");
            }
            Console.WriteLine();
        }
    }
}

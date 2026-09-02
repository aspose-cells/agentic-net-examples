// Title: Add a leftmost row-number column to an Excel sheet and export it as CSV using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file, inserts a new column at the very beginning, writes sequential numbers (1,2,3…) into that column, and writes the result out as a CSV using Aspose.Cells. | Produce a .NET snippet that adds an index column to the left side of the first worksheet, populates it with row identifiers, and saves the workbook in CSV format. | Generate a program that numbers every populated row, places the numbers in column A, and converts the workbook to a CSV file with Aspose.Cells.
// Common Searches: Aspose.Cells how to insert a column at position 0 before CSV export in C# | C# convert Excel workbook to CSV while adding a leading row number column using Aspose.Cells | add sequential row index to Excel sheet and save as CSV with Aspose.Cells .NET | prepend row numbers column to first worksheet then export to CSV Aspose.Cells example | save xlsx as csv with custom leftmost column in C# Aspose.Cells
// Tags: insert leftmost column Aspose.Cells | sequential row numbering Excel C# | save as CSV after column insertion Aspose.Cells | prepend row index column .NET | row number column generation Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an existing workbook, inserts a new column at the very left, fills that column with sequential row numbers starting at 1 for all populated rows, and then saves the modified worksheet as a CSV file using Aspose.Cells for .NET.
class WorkbookToCsvWithRowNumbers
{
    static void Main()
    {
        // Paths for source workbook and destination CSV
        string sourcePath = "input.xlsx";
        string destPath = "output.csv";

        // Load the existing workbook
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert a new column at the very left (index 0)
        cells.InsertColumn(0);

        // Determine the number of rows that contain data (including header)
        int totalRows = cells.MaxDataRow + 1;

        // Populate the new column with sequential row numbers starting from 1
        for (int row = 0; row < totalRows; row++)
        {
            cells[row, 0].PutValue(row + 1);
        }

        // Save the modified workbook as CSV
        workbook.Save(destPath, SaveFormat.Csv);
    }
}

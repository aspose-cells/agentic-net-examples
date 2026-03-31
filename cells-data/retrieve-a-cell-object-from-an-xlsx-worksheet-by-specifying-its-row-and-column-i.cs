using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (you could also load an existing file with new Workbook("file.xlsx"))
        Workbook workbook = new Workbook();

        // Get the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Put some initial data so we can see the effect
        worksheet.Cells["A1"].PutValue("Original");

        // Retrieve the cell at the specified zero‑based row and column indices.
        // Example: row = 0 (first row), column = 1 (second column) corresponds to cell B1.
        Cell cell = worksheet.Cells[0, 1];   // using the Cells indexer

        // You can now work with the retrieved cell, e.g., set a value
        cell.PutValue("Retrieved");

        // Save the workbook to verify the changes
        workbook.Save("RetrievedCellDemo.xlsx");
    }
}
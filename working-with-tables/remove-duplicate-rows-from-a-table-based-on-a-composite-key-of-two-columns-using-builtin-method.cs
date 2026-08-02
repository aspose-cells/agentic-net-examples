// Title: C# – Remove duplicate rows in an Aspose.Cells worksheet using a composite key (ID & Name)
// Description: Demonstrates how to call Cells.RemoveDuplicates to eliminate rows that have identical values in the ID and Name columns of an Excel sheet, keep the first occurrence, handle header rows, and save the cleaned workbook as an XLSX file.
// Keywords: Aspose.Cells RemoveDuplicates C# | remove duplicate rows Excel Aspose | composite key duplicate removal | Aspose.Cells duplicate rows multiple columns | C# Excel duplicate elimination
// Common Searches: Aspose.Cells remove duplicate rows based on two columns | C# Cells.RemoveDuplicates composite key example | how to delete duplicate records in Excel using Aspose.Cells
// Developer Intent: Delete rows that share the same ID and Name values from an Excel worksheet with Aspose.Cells.
// Use Cases: Clean sales export files where the same customer ID and name appear more than once. | Prepare master contact lists by removing duplicate entries before reporting. | Consolidate inventory sheets by discarding rows with identical product code and description.
// AI Prompts: Write C# code that uses Aspose.Cells to remove duplicate rows based on three columns (e.g., ID, Name, Date). | Explain each parameter of Cells.RemoveDuplicates, focusing on the hasHeaders flag and columnOffsets array. | Show how to keep the last duplicate instead of the first when using Aspose.Cells RemoveDuplicates.

using System;
using Aspose.Cells;

namespace AsposeCellsRemoveDuplicatesDemo
{
    // Demonstrates how to call Cells.RemoveDuplicates to eliminate rows that have identical values in the ID and Name columns of an Excel sheet, keep the first occurrence, handle header rows, and save the cleaned workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Populate sample data -----
            // Header row
            cells[0, 0].PutValue("ID");
            cells[0, 1].PutValue("Name");
            cells[0, 2].PutValue("Value");

            // Data rows (some rows have duplicate composite keys: ID + Name)
            cells[1, 0].PutValue(1);   cells[1, 1].PutValue("John");   cells[1, 2].PutValue(100);
            cells[2, 0].PutValue(2);   cells[2, 1].PutValue("Jane");   cells[2, 2].PutValue(200);
            cells[3, 0].PutValue(1);   cells[3, 1].PutValue("John");   cells[3, 2].PutValue(150); // duplicate key
            cells[4, 0].PutValue(3);   cells[4, 1].PutValue("Bob");    cells[4, 2].PutValue(300);

            // ----- Define the range for duplicate removal -----
            int startRow = 0;          // include header
            int startColumn = 0;
            int endRow = 4;            // last data row index
            int endColumn = 2;         // last column index
            bool hasHeaders = true;    // first row contains column names

            // Columns that constitute the composite key (ID column and Name column)
            int[] columnOffsets = new int[] { 0, 1 };

            // ----- Remove duplicate rows based on the composite key -----
            cells.RemoveDuplicates(startRow, startColumn, endRow, endColumn, hasHeaders, columnOffsets);

            // ----- Save the result -----
            workbook.Save("RemoveDuplicatesCompositeKey.xlsx");
        }
    }
}

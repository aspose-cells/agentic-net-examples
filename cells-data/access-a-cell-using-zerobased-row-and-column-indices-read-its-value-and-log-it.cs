// Title: Read and log a specific cell value using zero‑based row and column indices with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, accesses a cell using zero‑based row and column numbers, and prints the cell's value to the console. | Show how to retrieve a cell's value from the first worksheet of a workbook by providing numeric row and column indices in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# read cell by numeric row and column index | How to get Excel cell value using zero based indices in .NET | C# example for logging a cell value from a workbook with Aspose.Cells | Retrieve cell content from specific row and column using Aspose.Cells API | Read value at row 3 column 2 in Excel using Aspose.Cells for .NET
// Tags: zero-based cell indexing Aspose.Cells | fetch cell value using row column indices Aspose.Cells | load workbook and access first worksheet C# | log cell content to console .NET | numeric coordinates cell access Aspose.Cells

using System;
using Aspose.Cells;

// The example loads 'input.xlsx', selects the first worksheet, reads the value of the cell at zero‑based row index 2 and column index 1, and writes the retrieved value to the console.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        string filePath = "input.xlsx";
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Zero‑based row and column indices
        int rowIndex = 2;      // third row (zero‑based)
        int columnIndex = 1;   // second column (zero‑based)

        // Read the cell value
        object cellValue = sheet.Cells[rowIndex, columnIndex].Value;

        // Log the value
        Console.WriteLine($"Value at row {rowIndex}, column {columnIndex}: {cellValue}");
    }
}

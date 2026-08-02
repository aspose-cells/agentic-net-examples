// Title: Import an array and fetch a specific element by index with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, import a string array vertically using Cells.ImportArray, retrieve the third item with the Excel INDEX function (=INDEX(A:A,3)), calculate the formula, read the same value via the Cells[row, column] indexer, and save the result to an .xlsx file.
// Keywords: Aspose.Cells | C# | .NET | ImportArray | Excel INDEX function | retrieve array element | cell indexer | calculate formula | smart markers
// Common Searches: Aspose.Cells import string array C# | How to get nth element from imported column Aspose.Cells | Use INDEX formula with Aspose.Cells workbook | Access cell by row and column after ImportArray | Calculate formulas in Aspose.Cells C#
// Developer Intent: Extract a particular element from an array that has been imported into an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Load a product list into Excel and display the third product using a formula. | Read a specific row value directly after bulk data import without looping. | Create a report that isolates selected items from a larger dataset imported via ImportArray.
// AI Prompts: Show C# code that imports a string array into a worksheet with Aspose.Cells and returns the element at a given zero‑based index using the INDEX function. | Provide an example of using the INDEX formula in Aspose.Cells to fetch the nth element of an imported column and then read the computed value programmatically. | Explain how to calculate formulas and access the result after importing an array with Aspose.Cells in a .NET application.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, import a string array vertically using Cells.ImportArray, retrieve the third item with the Excel INDEX function (=INDEX(A:A,3)), calculate the formula, read the same value via the Cells[row, column] indexer, and save the result to an .xlsx file.
class ImportSpecificArrayElementDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define an array of item names.
        string[] itemNames = new string[] { "Apple", "Banana", "Cherry", "Date", "Elderberry" };

        // Import the array vertically starting at cell A1 (row 0, column 0).
        cells.ImportArray(itemNames, 0, 0, true);

        // Set a formula that retrieves the third element (index 2) of the array.
        // This mimics the custom syntax &=Orders[2].ItemName.
        // In Excel the equivalent is =INDEX(A:A,3) because Excel uses 1‑based indexing.
        cells["B1"].Formula = "=INDEX(A:A,3)";

        // Calculate the formula so the result is stored in the cell.
        workbook.CalculateFormula();

        // Output the value obtained via the formula.
        Console.WriteLine("Third item via formula: " + cells["B1"].StringValue);

        // Directly access the third element from the imported array using the Cells indexer.
        // Row index 2 (third row), column index 0 (first column).
        string directValue = cells[2, 0].StringValue;
        Console.WriteLine("Third item via direct access: " + directValue);

        // Save the workbook to a file.
        workbook.Save("ImportSpecificArrayElementDemo.xlsx");
    }
}

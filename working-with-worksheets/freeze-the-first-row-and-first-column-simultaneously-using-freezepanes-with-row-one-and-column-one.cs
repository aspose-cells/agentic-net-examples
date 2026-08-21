// Title: Freeze First Row and First Column Simultaneously with Aspose.Cells FreezePanes (C#)
// Description: Creates a new Workbook, accesses the first Worksheet, applies Worksheet.FreezePanes(1, 1, 1, 1) to lock the top row and leftmost column, and saves the file as FreezeFirstRowAndColumn.xlsx.
// Keywords: Aspose.Cells | FreezePanes | C# | freeze first row | freeze first column | freeze top row and column | Excel pane freezing .NET | worksheet FreezePanes example
// Common Searches: Aspose.Cells freeze first row and column C# | How to use FreezePanes in Aspose.Cells .NET | Freeze top row and left column Excel with Aspose.Cells | C# code to lock header row and identifier column | Aspose.Cells FreezePanes parameters
// Developer Intent: Lock the worksheet’s header row and identifier column in a single operation.
// Use Cases: Scroll large data tables while keeping header rows visible | Design a data‑entry template with fixed row and column labels | Generate automated reports where headings stay static during navigation | Build dashboards where the first row/column act as persistent filters
// AI Prompts: Provide a C# example to freeze multiple rows and columns using Aspose.Cells FreezePanes with custom indices. | Show code that freezes the first two rows and the first three columns in an Aspose.Cells workbook. | Explain each parameter of Worksheet.FreezePanes and how zero‑based indexing influences the frozen area.

using System;
using Aspose.Cells;

namespace FreezeFirstRowAndColumn
{
    // Creates a new Workbook, accesses the first Worksheet, applies Worksheet.FreezePanes(1, 1, 1, 1) to lock the top row and leftmost column, and saves the file as FreezeFirstRowAndColumn.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the first row and first column.
            // Parameters: row index, column index, number of frozen rows, number of frozen columns.
            // Row and column indices are zero‑based, so index 1 refers to the second row/column.
            // Freezing 1 row and 1 column will lock the top row (A1) and leftmost column (A column).
            worksheet.FreezePanes(1, 1, 1, 1);

            // Save the workbook to a file
            workbook.Save("FreezeFirstRowAndColumn.xlsx");
        }
    }
}

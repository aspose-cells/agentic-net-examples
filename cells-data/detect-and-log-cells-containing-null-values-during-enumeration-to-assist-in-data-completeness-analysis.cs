// Title: Log null or missing cells while enumerating the used range of an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to iterate over the used range of a worksheet and prints the address of every cell whose Value property is null. | Create a C# routine that employs Row.GetCellOrNull and Cells.CheckCell to identify both null-valued and non‑instantiated cells and logs their addresses. | Generate a C# example that gathers all null or missing cell addresses into a collection and exports the list to a CSV file using Aspose.Cells.
// Common Searches: Aspose.Cells C# find empty cells in used range of worksheet | how to detect null values in Excel cells using Aspose.Cells .NET | enumerate worksheet rows and columns and log missing cells with Aspose.Cells API | C# Aspose.Cells check for cells that were never created in a workbook
// Tags: null cell detection Aspose.Cells C# | enumerate used range cells Aspose.Cells | Row.GetCellOrNull usage | Cells.CheckCell missing cell detection | log empty Excel cells Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsNullDetectionDemo
{
    // Demonstrates creating a workbook, inserting intentional empty and null cells, then scanning the used range with Row.GetCellOrNull and Cells.CheckCell to log both null-valued and non‑instantiated cells, and finally saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data with intentional null/empty cells
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");

            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            // C2 left empty (null)

            cells["A3"].PutValue(2);
            // B3 left empty (null)
            cells["C3"].PutValue(85);

            cells["A4"].PutValue(3);
            cells["B4"].PutValue("Charlie");
            cells["C4"].PutValue(null); // Explicit null assignment

            // Determine the used range boundaries
            int maxRow = cells.MaxDataRow;      // zero‑based index of last row with data
            int maxCol = cells.MaxDataColumn;   // zero‑based index of last column with data

            Console.WriteLine("Scanning for cells with null values...");

            // Iterate through each row in the used range
            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
            {
                // Obtain the Row object
                Row row = cells.Rows[rowIndex];

                // Iterate through each column in the used range
                for (int colIndex = 0; colIndex <= maxCol; colIndex++)
                {
                    // Use Row.GetCellOrNull to safely retrieve the cell (may return null if cell not instantiated)
                    Cell cell = row.GetCellOrNull(colIndex);

                    // If the cell exists, check its Value property for null
                    if (cell != null && cell.Value == null)
                    {
                        // Log the address of the cell containing a null value
                        Console.WriteLine($"Null value found at {cell.Name} (Row {rowIndex + 1}, Column {colIndex + 1})");
                    }
                    // If the cell object itself is null, it means the cell was never created (also considered empty)
                    else if (cell == null)
                    {
                        // Use CheckCell to confirm the cell truly does not exist
                        Cell checkCell = cells.CheckCell(rowIndex, colIndex);
                        if (checkCell == null)
                        {
                            // Log the address of the missing cell as a null/empty entry
                            string address = CellsHelper.CellIndexToName(rowIndex, colIndex);
                            Console.WriteLine($"Missing (null) cell at {address} (Row {rowIndex + 1}, Column {colIndex + 1})");
                        }
                    }
                }
            }

            // Save the workbook (demonstrates lifecycle compliance)
            workbook.Save("NullDetectionDemo.xlsx");
        }
    }
}

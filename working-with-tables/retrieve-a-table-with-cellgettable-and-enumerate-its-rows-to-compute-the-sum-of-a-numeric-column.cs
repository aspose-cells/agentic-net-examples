// Title: C# Aspose.Cells – Retrieve a Table with Cell.GetTable and Sum a Numeric Column
// Description: Demonstrates creating an Excel table, obtaining its ListObject via Cell.GetTable, iterating data rows, converting values from the numeric column, and calculating the total. The sum is printed and the workbook saved.
// Keywords: Aspose.Cells | C# | .NET | Cell.GetTable | ListObject | Excel table | sum column | aggregate numeric values | enumerate table rows | retrieve table from cell | calculate column total
// Common Searches: Aspose.Cells get table from cell C# | How to sum a column in an Aspose.Cells ListObject | Cell.GetTable example for .NET | Iterate Aspose.Cells table rows and calculate total | C# aggregate numeric column in Excel table using Aspose
// Developer Intent: Obtain a ListObject via Cell.GetTable and compute the sum of its numeric column.
// Use Cases: Calculate total sales amount from a worksheet table for financial reports. | Aggregate inventory quantities across product rows to determine overall stock. | Sum student test scores stored in a table to generate class averages.
// AI Prompts: Generate C# code that uses Aspose.Cells Cell.GetTable to retrieve a table and sum a specified numeric column, handling int and double values. | Explain step‑by‑step how to loop through the rows of a ListObject obtained with GetTable and compute a column total, including non‑numeric cell handling. | Create a reusable C# method that accepts a Worksheet and a column name, returns the sum of that column using Aspose.Cells, and demonstrates its usage.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates creating an Excel table, obtaining its ListObject via Cell.GetTable, iterating data rows, converting values from the numeric column, and calculating the total. The sum is printed and the workbook saved.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data with a header row and a numeric column
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("C");
        cells["B4"].PutValue(30);

        // Create a table that includes the data range (A1:B4)
        int tableIdx = worksheet.ListObjects.Add("A1", "B4", true);
        ListObject table = worksheet.ListObjects[tableIdx];

        // Retrieve the same table using a cell that belongs to it
        Cell sampleCell = cells["A2"];               // any cell inside the table
        ListObject retrievedTable = sampleCell.GetTable();

        // Compute the sum of the numeric column (second column, index 1)
        double sum = 0;
        int dataStartRow = retrievedTable.StartRow + 1; // skip header row
        int dataEndRow = retrievedTable.EndRow;
        int numericColIndex = retrievedTable.StartColumn + 1; // second column in the table

        for (int row = dataStartRow; row <= dataEndRow; row++)
        {
            object val = cells[row, numericColIndex].Value;
            if (val is double d)
                sum += d;
            else if (val is int i)
                sum += i;
            else if (double.TryParse(val?.ToString(), out double parsed))
                sum += parsed;
        }

        // Output the computed sum
        Console.WriteLine($"Sum of column '{table.ListColumns[1].Name}': {sum}");

        // Save the workbook (optional)
        workbook.Save("TableSumDemo.xlsx");
    }
}

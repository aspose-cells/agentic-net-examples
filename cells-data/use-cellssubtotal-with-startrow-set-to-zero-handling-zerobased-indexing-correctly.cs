// Title: Create a sum subtotal in an Excel worksheet using Aspose.Cells C# with zero‑based start row indexing
// AI Prompts: Generate C# code that builds a workbook, sets up a header row, populates data, defines a CellArea beginning at row 0, and calls Cells.Subtotal to group by the first column and sum the third column. | Demonstrate how to configure Cells.Subtotal in Aspose.Cells when the start row is zero, including the correct zero‑based EndRow calculation. | Provide a step‑by‑step example of applying a subtotal to the Sales column while grouping by Region using Aspose.Cells with zero‑based row numbers.
// Common Searches: Aspose.Cells C# subtotal startrow zero based indexing example | How to use Cells.Subtotal with zero‑based row numbers in .NET | C# code to group rows by Region and sum Sales using Aspose.Cells Subtotal | Define CellArea that includes header row for subtotal in Aspose.Cells | Apply sum subtotal on a specific column with Aspose.Cells zero‑based rows
// Tags: Aspose.Cells Subtotal zero-based indexing | C# define CellArea for subtotal | Aspose.Cells sum subtotal by column | Excel workbook subtotal using Aspose.Cells C# | zero-based row indexing Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, adds a header row and sample data, defines a CellArea that starts at row 0 to include the header, and uses Cells.Subtotal to group by the Region column (first column) while summing the Sales column (third column) with zero‑based indexing. The workbook is saved as SubtotalZeroBasedDemo.xlsx.
public class SubtotalZeroBasedDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add header row (zero‑based row index 0)
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            // Sample data (5 rows)
            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 }
            };

            // Populate data starting at row index 1 (zero‑based)
            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Region
                cells[i + 1, 1].PutValue(data[i, 1]); // Product
                cells[i + 1, 2].PutValue(data[i, 2]); // Sales
            }

            // Define the cell area covering header + data.
            // StartRow = 0 (first row), EndRow = 5 (last data row index)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0), // 5 (zero‑based index of last row)
                EndColumn = 2
            };

            // Apply subtotal:
            // - groupBy: column 0 (Region) – zero‑based index
            // - function: Sum
            // - totalList: column 2 (Sales) – zero‑based index
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 });

            // Save the workbook
            workbook.Save("SubtotalZeroBasedDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        SubtotalZeroBasedDemo.Run();
    }
}

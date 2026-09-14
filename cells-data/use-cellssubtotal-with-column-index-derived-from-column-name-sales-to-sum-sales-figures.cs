// Title: Use Aspose.Cells Subtotal in C# to sum a dynamically located "Sales" column grouped by Region
// AI Prompts: Generate C# code that scans the first worksheet row to locate the column named "Sales" and then calls Worksheet.Cells.Subtotal with ConsolidationFunction.Sum, grouping by the Region column. | Demonstrate how to build a CellArea covering the data range and invoke Subtotal using a column index discovered at runtime to aggregate sales values in an Aspose.Cells workbook. | Provide a complete .NET example that groups rows by the first column and totals the dynamically identified sales column with Aspose.Cells Subtotal, then saves the workbook.
// Common Searches: aspnet find column index by header name and apply subtotal with Aspose.Cells | c# Aspose.Cells subtotal sum sales column determined at runtime | how to group rows by region and total sales using Aspose.Cells Subtotal method | using Aspose.Cells Subtotal with dynamic column index in .NET Core
// Tags: Aspose.Cells Subtotal with dynamic column index | C# locate header index Aspose.Cells | sum sales column Aspose.Cells | group rows by region Aspose.Cells | CellArea definition for subtotal Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    // The example creates a workbook, inserts sample data with a "Sales" header, scans the header row to determine the zero‑based index of the "Sales" column, defines a CellArea that spans the data range, and calls worksheet.Cells.Subtotal to group rows by the first column (Region) and sum the values in the identified sales column. The resulting file is saved as SubtotalBySalesColumn.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data with a header row that includes "Sales"
            // Header: Region | Product | Sales
            cells["A1"].PutValue("Region");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                { "North", "Widget", 5000 },
                { "North", "Gadget", 3000 },
                { "South", "Widget", 6000 },
                { "South", "Gadget", 4000 },
                { "West",  "Widget", 4500 }
            };

            // Fill the worksheet with the sample data starting from row 2 (zero‑based index 1)
            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Determine the zero‑based column index of the "Sales" header
            int salesColumnIndex = -1;
            int headerRow = 0;
            int totalColumns = cells.MaxColumn; // number of columns used

            for (int col = 0; col < totalColumns; col++)
            {
                if (cells[headerRow, col].StringValue.Equals("Sales", StringComparison.OrdinalIgnoreCase))
                {
                    salesColumnIndex = col;
                    break;
                }
            }

            if (salesColumnIndex == -1)
            {
                Console.WriteLine("Column \"Sales\" not found.");
                return;
            }

            // Define the cell area that contains the data (including the header row)
            // EndRow is the last row with data (header + data rows)
            int endRow = data.GetLength(0); // data rows count
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = endRow,
                EndColumn = totalColumns - 1
            };

            // Apply subtotal:
            // - Group by the first column (Region) -> index 0
            // - Use SUM function
            // - Subtotal the column identified as "Sales"
            worksheet.Cells.Subtotal(
                area,
                0, // group by column 0 (Region)
                ConsolidationFunction.Sum,
                new int[] { salesColumnIndex }
            );

            // Save the workbook
            workbook.Save("SubtotalBySalesColumn.xlsx");
        }
    }
}

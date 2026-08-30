// Title: Create CountNumbers subtotals in column K grouped by column A without outline levels using Aspose.Cells for .NET
// AI Prompts: Generate C# code that applies the numeric count consolidation function to column K, groups rows by column A, and disables outline groups using Aspose.Cells. | Demonstrate how to configure the Subtotal method parameters to prevent outline hierarchy and save the workbook as an .xlsx file.
// Common Searches: how to add count numbers subtotal to a specific column using Aspose.Cells C# | aspocells subtotal method without creating outline groups | c# group rows by column A and count numeric values in column K with Aspose.Cells | prevent outline groups when applying subtotal in Aspose.Cells workbook | how to use CountNums function for subtotals in Aspose.Cells
// Tags: Aspose.Cells ConsolidationFunction.CountNums subtotal column K C# | Aspose.Cells subtotal without outline hierarchy | remove outline hierarchy from Excel subtotals Aspose.Cells | C# add numeric count subtotals to worksheet | Excel workbook subtotal example using Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // The program creates a workbook, fills columns A and K with sample data, adds a CountNumbers subtotal on column K grouped by column A while suppressing outline levels, and saves the file as Subtotal_CountNums_ColumnK.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Sample data setup (columns A to K, rows 1 to 6)
            // ------------------------------------------------------------
            // Header row
            cells["A1"].PutValue("Group");
            cells["K1"].PutValue("Values"); // Column K (index 10)

            // Data rows
            object[,] data = new object[,]
            {
                { "A", 12 },
                { "A", 7  },
                { "B", 5  },
                { "B", 9  },
                { "C", 3  },
                { "C", 8  }
            };

            // Populate the sample data (starting from row 2)
            for (int i = 0; i < data.GetLength(0); i++)
            {
                // Column A (index 0) – group identifier
                cells[i + 1, 0].PutValue(data[i, 0]);
                // Column K (index 10) – numeric values to be counted
                cells[i + 1, 10].PutValue(data[i, 1]);
            }

            // ------------------------------------------------------------
            // Define the range that contains the data (A1:K6)
            // ------------------------------------------------------------
            CellArea area = CellArea.CreateCellArea("A1", "K6");

            // ------------------------------------------------------------
            // Add subtotals:
            //   - Group by column A (zero‑based index 0)
            //   - Use CountNums function (counts numeric values)
            //   - Apply subtotal to column K (zero‑based index 10)
            //   - Do not replace existing subtotals, no page breaks,
            //     and place summary below data (set to false to avoid
            //     outline levels)
            // ------------------------------------------------------------
            worksheet.Cells.Subtotal(
                area,                     // range
                0,                        // group by column A
                ConsolidationFunction.CountNums, // CountNumbers function
                new int[] { 10 },         // subtotal column K
                false,                    // replace existing subtotals
                false,                    // no page breaks between groups
                false                     // summary not below data (prevents outline)
            );

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("Subtotal_CountNums_ColumnK.xlsx");
        }
    }
}

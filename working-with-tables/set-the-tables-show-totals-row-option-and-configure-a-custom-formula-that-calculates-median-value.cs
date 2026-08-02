// Title: Aspose.Cells .NET – Show Table Totals Row and Apply a Custom Median Formula
// Description: Creates a workbook, adds a ListObject (A1:B5), enables the totals row, names the numeric column, sets TotalsCalculation to Custom, and inserts a median formula (=MEDIAN([Value])) using SetCustomTotalsRowFormula.
// Keywords: Aspose.Cells | C# | Excel table totals row | custom totals formula | median calculation | SetCustomTotalsRowFormula | ListObject | ListColumn | show totals row | Aspose.Cells .NET example
// Common Searches: Aspose.Cells show totals row in table | C# set custom totals row formula median | How to calculate median in Aspose.Cells table | SetCustomTotalsRowFormula example | Enable totals row for ListObject Aspose.Cells
// Developer Intent: Enable a table's totals row and compute the column median with a custom formula in Aspose.Cells for .NET.
// Use Cases: Add a totals row to a sales table that automatically displays the median revenue. | Generate a dynamic report where the median of measurement data updates as values change. | Create Excel workbooks that summarize data with custom statistical calculations like median.
// AI Prompts: Write C# code using Aspose.Cells to add a ListObject, show its totals row, and set a custom median formula for a column. | Explain the parameters of SetCustomTotalsRowFormula and how to reference a column name in the formula. | Show how to modify the custom totals row formula at runtime based on user‑selected aggregation (median, average, sum).

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a ListObject (A1:B5), enables the totals row, names the numeric column, sets TotalsCalculation to Custom, and inserts a median formula (=MEDIAN([Value])) using SetCustomTotalsRowFormula.
    public class TableTotalsMedianDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data with a header and some numeric values
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("C");
                cells["B4"].PutValue(30);
                cells["A5"].PutValue("D");
                cells["B5"].PutValue(40);

                // Add a table that includes the data range (A1:B5)
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Enable the totals row for the table
                table.ShowTotals = true;

                // Access the second column ("Value") in the table
                ListColumn valueColumn = table.ListColumns[1];
                // Optionally set a friendly name for the column (used in the formula)
                valueColumn.Name = "Value";

                // Set the totals calculation type to Custom
                valueColumn.TotalsCalculation = TotalsCalculation.Custom;

                // Define a custom formula that calculates the median of the column values
                // The formula uses the column name within square brackets.
                // isR1C1 = false (A1 style), isLocal = false (invariant culture)
                valueColumn.SetCustomTotalsRowFormula("=MEDIAN([Value])", false, false);

                // Save the workbook to a file
                string outputPath = "TableTotalsMedianDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableTotalsMedianDemo.Run();
        }
    }
}

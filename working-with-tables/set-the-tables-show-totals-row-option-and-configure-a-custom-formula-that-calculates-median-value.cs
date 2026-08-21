// Title: Aspose.Cells .NET – Show Table Totals Row and Apply a Custom Median Formula (C#)
// Description: Creates a workbook, adds a ListObject covering A1:B4, enables the totals row, names the numeric column, sets its TotalsCalculation to Custom, and assigns the formula "=MEDIAN([Value])" to compute the column median before saving the file.
// Keywords: Aspose.Cells | .NET | C# | ListObject | ShowTotals | totals row | custom totals formula | median calculation | SetCustomTotalsRowFormula | Excel table median | Aspose.Cells example
// Common Searches: how to enable totals row in Aspose.Cells | Aspose.Cells custom totals row formula | calculate median in table column Aspose.Cells | SetCustomTotalsRowFormula C# example | Aspose.Cells ListObject median | show totals row for table Aspose.Cells .NET
// Developer Intent: Add a table, turn on its totals row, and define a custom median calculation for a column.
// Use Cases: Financial reports that need the median expense displayed in the totals row. | Statistical dashboards summarizing median values of measurement data. | Automated Excel generation where dynamic data sets require a median summary. | Quality‑control sheets that show median defect rates directly in the table footer.
// AI Prompts: Write C# code using Aspose.Cells to create a table, enable the totals row, and set a custom median formula for the "Value" column. | Explain how SetCustomTotalsRowFormula works with column name references in Aspose.Cells and demonstrate a median calculation. | Provide a step‑by‑step guide to add a ListObject, show its totals row, and apply any custom formula (e.g., =MEDIAN([Column])) in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a ListObject covering A1:B4, enables the totals row, names the numeric column, sets its TotalsCalculation to Custom, and assigns the formula "=MEDIAN([Value])" to compute the column median before saving the file.
    public class TableMedianTotalsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (header + three rows)
                worksheet.Cells["A1"].PutValue("Item");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B4"].PutValue(30);

                // Add a table that covers the data range A1:B4
                int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Enable the totals row for the table
                table.ShowTotals = true;

                // Access the second column ("Value") and set its name (optional but makes the formula clearer)
                ListColumn valueColumn = table.ListColumns[1];
                valueColumn.Name = "Value";

                // Specify that the totals calculation for this column is custom
                valueColumn.TotalsCalculation = TotalsCalculation.Custom;

                // Set a custom formula that calculates the median of the column values
                // The formula uses the column name inside square brackets, e.g., =MEDIAN([Value])
                valueColumn.SetCustomTotalsRowFormula("=MEDIAN([Value])", false, false);

                // Save the workbook to a file
                workbook.Save("TableMedianTotalsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableMedianTotalsDemo.Run();
        }
    }
}

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
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

                // Populate sample data (header + numeric values)
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Quantity");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Orange");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("Banana");
                cells["B4"].PutValue(30);
                cells["A5"].PutValue("Grape");
                cells["B5"].PutValue(40);

                // Add a table that includes the data range (A1:B5)
                int tableIndex = worksheet.ListObjects.Add("A1", "B5", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Show the totals row for the table
                table.ShowTotals = true;

                // Configure the second column (Quantity) to use a custom totals calculation
                ListColumn quantityColumn = table.ListColumns[1]; // zero‑based index
                quantityColumn.TotalsCalculation = TotalsCalculation.Custom;

                // Set a custom formula that calculates the median of the Quantity column
                // The formula uses structured table reference syntax: =MEDIAN([Quantity])
                quantityColumn.SetCustomTotalsRowFormula("=MEDIAN([Quantity])", false, false);

                // Optionally, set a label for the totals row in the first column
                table.ListColumns[0].TotalsRowLabel = "Median";

                // Save the workbook
                string outputPath = "TableTotalsMedianDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TableTotalsMedianDemo.Run();
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class SetCustomAverageInTotalsRow
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
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(30);

                // Add a table that includes the data range (A1:B4) and enable the totals row
                int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.ShowTotals = true;

                // The column we want the average for is the second column (Quantity)
                ListColumn quantityColumn = table.ListColumns[1];

                // Set the totals calculation type to Custom
                quantityColumn.TotalsCalculation = TotalsCalculation.Custom;

                // Define a custom formula for the totals row using structured reference syntax
                quantityColumn.SetCustomTotalsRowFormula("=AVERAGE([Quantity])", false, false);

                // Save the workbook
                string outputPath = "CustomAverageTotalsRow.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetCustomAverageInTotalsRow.Run();
        }
    }
}
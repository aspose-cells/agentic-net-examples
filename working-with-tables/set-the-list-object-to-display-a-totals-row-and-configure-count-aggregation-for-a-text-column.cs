using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace Demo
{
    class ListObjectTotalsCountDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (text column + numeric column)
                worksheet.Cells["A1"].PutValue("Item");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("Apple");
                worksheet.Cells["B4"].PutValue(15);

                // Add a ListObject (table) that includes the data range
                int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Enable the totals row
                table.ShowTotals = true;

                // Set count aggregation for the text column (first column)
                table.ListColumns[0].TotalsCalculation = TotalsCalculation.Count;

                // Optional: set a label for the totals row of that column
                table.ListColumns[0].TotalsRowLabel = "Count";

                // Save the workbook
                string outputPath = "ListObjectTotalsCountDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ListObjectTotalsCountDemo.Run();
        }
    }
}
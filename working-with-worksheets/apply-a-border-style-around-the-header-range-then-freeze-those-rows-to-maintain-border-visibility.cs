using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHeaderBorderAndFreeze
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate header row (A1 to D1) and some sample data
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["C1"].PutValue("Category");
                cells["D1"].PutValue("Price");

                for (int i = 2; i <= 10; i++)
                {
                    cells[$"A{i}"].PutValue(i - 1);
                    cells[$"B{i}"].PutValue($"Item {i - 1}");
                    cells[$"C{i}"].PutValue("General");
                    cells[$"D{i}"].PutValue(10.0 * (i - 1));
                }

                // Define the header range (first row)
                AsposeRange headerRange = cells.CreateRange("A1:D1");

                // Apply a thick black outline border around the header range
                headerRange.SetOutlineBorders(CellBorderType.Thick, Color.Black);

                // Freeze the header row so it stays visible while scrolling
                // Freeze at cell A2 with 1 frozen row and 0 frozen columns
                worksheet.FreezePanes("A2", 1, 0);

                // Save the workbook
                string outputPath = "HeaderBorderAndFreeze.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
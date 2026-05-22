using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetCalculateAfterDeletionDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column A (rows 1-5)
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["A4"].PutValue(40);
            cells["A5"].PutValue(50);

            // Add a formula that sums the values in column A
            cells["B1"].Formula = "=SUM(A1:A5)";

            // Initial calculation to evaluate the formula
            workbook.CalculateFormula();

            Console.WriteLine("Initial sum (B1): " + cells["B1"].StringValue); // Expected 150

            // Delete the second row (index 1) and recalculate
            cells.DeleteRow(1);               // Removes row 2 (value 20)
            workbook.CalculateFormula();     // Force recalculation after deletion

            Console.WriteLine("After deleting row 2, sum (B1): " + cells["B1"].StringValue); // Expected 130

            // Delete two rows starting from the third row (current indices after previous deletion)
            cells.DeleteRows(2, 2);           // Removes rows that originally were 4 and 5 (values 40 and 50)
            workbook.CalculateFormula();     // Force recalculation after second deletion

            Console.WriteLine("After deleting rows 4-5, sum (B1): " + cells["B1"].StringValue); // Expected 40 (10+30)

            // Save the workbook to verify the final state
            string outputPath = "WorksheetCalculateAfterDeletionDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
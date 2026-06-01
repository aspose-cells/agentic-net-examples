using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class VerifyCellContentLengthAfterCompatibilityToggle
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare source data with a string longer than 255 characters
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header
                dataSheet.Cells["A1"].Value = "Category";
                dataSheet.Cells["B1"].Value = "LongText";

                // Sample data row
                dataSheet.Cells["A2"].Value = "Item1";

                // Create a long string (300 characters)
                string longString = new string('x', 300);
                dataSheet.Cells["B2"].Value = longString;

                // -------------------------------------------------
                // 2. Add a pivot table based on the source data
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
                // Add pivot table: source range A1:B2, destination start cell A4
                int pivotIndex = pivotSheet.PivotTables.Add("PivotTable", "A1:B2", "A4", false);
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add fields: Category as Row, LongText as Data (to display the text)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // LongText column

                // -------------------------------------------------
                // 3. First refresh with default compatibility (true)
                //    This will truncate the long string to 255 characters
                // -------------------------------------------------
                pivotTable.IsExcel2003Compatible = true; // explicit for clarity
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Locate the cell that contains the pivot data (first data cell)
                // After refresh, data starts at the cell right below the row field header.
                // Row field header is at A4, data starts at B5 (zero‑based indices)
                Cell pivotDataCell = pivotSheet.Cells[4, 1]; // B5

                string truncatedValue = pivotDataCell.StringValue;
                Console.WriteLine("After first refresh (compatible = true):");
                Console.WriteLine($"Length of pivot cell value: {truncatedValue.Length}");
                Console.WriteLine($"Is truncated (length < original): {truncatedValue.Length < longString.Length}");

                // -------------------------------------------------
                // 4. Disable Excel 2003 compatibility and refresh again
                //    The long string should now appear in full (no truncation)
                // -------------------------------------------------
                pivotTable.IsExcel2003Compatible = false;
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Read the value again from the same cell location
                string fullValue = pivotSheet.Cells[4, 1].StringValue;
                Console.WriteLine("\nAfter second refresh (compatible = false):");
                Console.WriteLine($"Length of pivot cell value: {fullValue.Length}");
                Console.WriteLine($"Matches original length: {fullValue.Length == longString.Length}");

                // -------------------------------------------------
                // 5. Verify that the content length remains unchanged after disabling compatibility
                // -------------------------------------------------
                bool lengthUnchanged = fullValue.Length == longString.Length;
                Console.WriteLine($"\nVerification result – length unchanged: {lengthUnchanged}");

                // Save the workbook (using the standard save method)
                string outputPath = "VerifyCellContentLengthAfterCompatibilityToggle.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyCellContentLengthAfterCompatibilityToggle.Run();
        }
    }
}
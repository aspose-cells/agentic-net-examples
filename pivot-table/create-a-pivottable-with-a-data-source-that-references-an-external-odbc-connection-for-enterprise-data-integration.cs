using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ------------------------------------------------------------
            // Prepare sample data for the PivotTable
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "PivotSheet";

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Sales");

            // Sample data rows
            string[] categories = { "Beverages", "Beverages", "Condiments", "Condiments", "Confections" };
            double[] sales = { 1200.5, 850.75, 430.0, 610.25, 980.0 };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(categories[i]); // Column A
                sheet.Cells[i + 2, 1].PutValue(sales[i]);     // Column B
            }

            // ------------------------------------------------------------
            // Create a PivotTable based on the sample data range
            // ------------------------------------------------------------
            // The source data range includes the header row and all data rows.
            string sourceData = $"A1:B{categories.Length + 1}";
            int pivotIndex = sheet.PivotTables.Add(sourceData, "D1", "EnterprisePivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // ------------------------------------------------------------
            // Configure the PivotTable fields
            // ------------------------------------------------------------
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            string outputPath = "EnterprisePivot.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
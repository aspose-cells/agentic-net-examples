// Title: Copy a Worksheet with Pivot Tables Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data and a pivot table, duplicate the worksheet with AddCopy, rename the copy, refresh its pivot tables, and save the file, ensuring the pivot configuration is preserved in the cloned sheet.
// Keywords: Aspose.Cells copy worksheet C# | duplicate sheet preserve pivot tables | AddCopy refresh pivot Aspose.Cells | clone worksheet with pivot .NET | Aspose.Cells pivot table copy example
// Common Searches: How to copy a worksheet and keep pivot tables in Aspose.Cells | Refresh pivot tables after duplicating a sheet in C# | Copy worksheet with pivot tables using Aspose.Cells for .NET | Clone sheet with pivot layout Aspose.Cells
// Developer Intent: Create a copy of an existing worksheet while retaining and updating its pivot tables.
// Use Cases: Generate a backup analysis sheet that mirrors the original pivot layout. | Create multiple scenario worksheets that share the same pivot structure for comparative reporting. | Build a template workbook where pivot tables are duplicated across regional sheets.
// AI Prompts: Provide C# code that copies a worksheet containing pivot tables with Aspose.Cells and automatically refreshes the cloned pivots. | Explain how to duplicate a sheet with several pivot tables in Aspose.Cells and adjust their data sources after copying. | Show an example of using AddCopy and RefreshPivotTables to retain pivot configurations when cloning a worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data and a pivot table, duplicate the worksheet with AddCopy, rename the copy, refresh its pivot tables, and save the file, ensuring the pivot configuration is preserved in the cloned sheet.
    public class DuplicateWorksheetWithPivot
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and give it a meaningful name
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate sample data for the pivot table
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Amount");
                sourceSheet.Cells["A2"].PutValue("Food");
                sourceSheet.Cells["B2"].PutValue(120);
                sourceSheet.Cells["A3"].PutValue("Drink");
                sourceSheet.Cells["B3"].PutValue(80);
                sourceSheet.Cells["A4"].PutValue("Food");
                sourceSheet.Cells["B4"].PutValue(150);
                sourceSheet.Cells["A5"].PutValue("Drink");
                sourceSheet.Cells["B5"].PutValue(70);

                // Add a pivot table to the source worksheet
                int pivotIndex = sourceSheet.PivotTables.Add("A1:B5", "D2", "SalesPivot");
                PivotTable pivot = sourceSheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.CalculateData(); // Populate the pivot table

                // Duplicate the worksheet (including the pivot table) using AddCopy
                int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Name);
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
                copiedSheet.Name = "CopiedData";

                // Refresh pivot tables in the copied sheet to reflect any changes
                copiedSheet.RefreshPivotTables();

                // Save the workbook
                workbook.Save("DuplicatedWorksheetWithPivot.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}

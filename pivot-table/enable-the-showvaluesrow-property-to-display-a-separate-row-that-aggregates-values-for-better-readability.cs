// Title: Enable ShowValuesRow to add a separate values row in an Aspose.Cells pivot table using C#
// AI Prompts: Write C# code that creates a workbook, populates sample data, adds a pivot table on range A1:B5, and sets pivotTable.ShowValuesRow = true with Aspose.Cells. | Generate a C# example that adds multiple data fields to an Aspose.Cells pivot table while keeping the values row visible. | Provide a C# snippet that formats the values row (font, background) after enabling ShowValuesRow and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# ShowValuesRow property example for pivot tables | how to display a separate values row in an Excel pivot table with Aspose.Cells | C# code to enable values row in Aspose.Cells pivot table and save as xlsx | Aspose.Cells pivot table ShowValuesRow true usage in .NET | sample Aspose.Cells program that adds a values row to a pivot table
// Tags: Aspose.Cells pivot table ShowValuesRow | C# create Excel pivot table Aspose.Cells | display separate values row Excel Aspose | Aspose.Cells refresh calculate pivot data | save workbook with pivot table .xlsx Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, inserting sample data, adding a pivot table on range A1:B5, enabling the ShowValuesRow property to display a distinct values row, refreshing and calculating the pivot data, and saving the result as PivotTableShowValuesRowDemo.xlsx using Aspose.Cells for .NET.
    public class PivotTableShowValuesRowDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                Cells cells = sheet.Cells;
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "Food";
                cells["B2"].Value = 120;
                cells["A3"].Value = "Food";
                cells["B3"].Value = 80;
                cells["A4"].Value = "Beverage";
                cells["B4"].Value = 150;
                cells["A5"].Value = "Beverage";
                cells["B5"].Value = 70;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category as row field, Amount as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Amount

                // Enable the ShowValuesRow property to display a separate values row
                pivotTable.ShowValuesRow = true;

                // Refresh and calculate the pivot data using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook to a file
                string outputPath = "PivotTableShowValuesRowDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}

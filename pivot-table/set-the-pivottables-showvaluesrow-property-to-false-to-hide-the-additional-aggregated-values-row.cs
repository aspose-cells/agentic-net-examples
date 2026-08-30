// Title: How to hide the aggregated values row in an Aspose.Cells PivotTable using C#
// AI Prompts: Create a new workbook, add sample data, build a PivotTable on range A1:B5, set the ShowValuesRow property to false, and save the file as XLSX with Aspose.Cells for .NET. | Programmatically configure row and data fields for a PivotTable and disable the extra values row in C# using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# hide values row in pivot table | Set ShowValuesRow false example Aspose.Cells .NET | Remove aggregated values row from Excel pivot using Aspose.Cells | How to disable the values row in a PivotTable with C# Aspose.Cells | C# code to hide extra totals row in Aspose.Cells pivot table
// Tags: Aspose.Cells PivotTable ShowValuesRow | C# hide pivot values row | Aspose.Cells disable aggregated row | Excel pivot hide values row .NET | Aspose.Cells set ShowValuesRow false

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills it with sample data, adds a PivotTable on A1:B5, assigns Category as rows and Amount as data, sets ShowValuesRow to false to hide the additional aggregated values row, and saves the workbook as PivotTableShowValuesRowDemo.xlsx.
    public class PivotTableShowValuesRowDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Pivot table created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
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

            // Add a pivot table (range A1:B5, destination C3)
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "C3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Amount

            // Hide the additional aggregated values row
            pivotTable.ShowValuesRow = false;

            // Save the workbook
            workbook.Save("PivotTableShowValuesRowDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}

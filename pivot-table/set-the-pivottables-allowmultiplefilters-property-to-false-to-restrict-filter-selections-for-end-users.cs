// Title: Aspose.Cells C# – Set PivotTable AllowMultipleFiltersPerField to False
// Description: Shows how to create a workbook, add sample data, build a PivotTable, assign row and data fields, and configure AllowMultipleFiltersPerField = false so users can pick only one filter item per field. The workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells | C# PivotTable | AllowMultipleFiltersPerField | disable multiple filters | single filter per field | Excel pivot settings | Aspose.Cells .NET example | pivot table filter restriction | workbook automation | data analysis Excel
// Common Searches: Aspose.Cells set AllowMultipleFiltersPerField false | C# pivot table limit filter selections | disable multiple filters in Aspose.Cells PivotTable | single filter per field Excel pivot Aspose | how to restrict pivot table filters using Aspose.Cells
// Developer Intent: Disable multiple filter selections per field in an Aspose.Cells PivotTable (C#).
// Use Cases: Create a sales dashboard where each category can be filtered by only one value, preventing overlapping selections. | Generate a financial report that enforces a single filter per field to maintain data consistency. | Build an Excel‑based analytics tool for end‑users that simplifies pivot interactions by allowing only one filter choice per dimension.
// AI Prompts: Write C# code with Aspose.Cells to add a PivotTable and set AllowMultipleFiltersPerField to false. | Explain the impact of AllowMultipleFiltersPerField on user experience in an Excel pivot table and provide a complete example. | Provide step‑by‑step instructions to configure a PivotTable in Aspose.Cells so each field permits only a single filter selection.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add sample data, build a PivotTable, assign row and data fields, and configure AllowMultipleFiltersPerField = false so users can pick only one filter item per field. The workbook is saved as an .xlsx file.
    public class PivotTableAllowMultipleFiltersDemo
    {
        public static void Main(string[] args)
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Food";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Drink";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Food";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Drink";
            cells["B5"].Value = 60;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Restrict filter selections: disallow multiple filters per field
            pivotTable.AllowMultipleFiltersPerField = false;

            // Determine output file path
            string outputPath = "PivotTable_AllowMultipleFiltersPerField_False.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}

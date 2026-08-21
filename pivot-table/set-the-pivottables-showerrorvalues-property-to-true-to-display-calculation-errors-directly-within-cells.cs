// Title: Aspose.Cells C# – Enable ShowErrorValues in a PivotTable to Display Calculation Errors
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable, define a calculated field that triggers a division‑by‑zero, and configure the PivotTable to show error values (using ShowErrorValues / DisplayErrorString) before refreshing and saving the file.
// Keywords: Aspose.Cells PivotTable ShowErrorValues | DisplayErrorString property | C# pivot table error handling | show #DIV/0! in Excel pivot | Aspose.Cells calculated field error | Excel error display .NET | pivot table data validation | Aspose.Cells C# example
// Common Searches: how to show error values in Aspose.Cells pivot table | Aspose.Cells DisplayErrorString true example | show #DIV/0! in pivot table using C# | Enable ShowErrorValues property Aspose.Cells | pivot table calculation error display .NET
// Developer Intent: Configure a PivotTable so that calculation errors appear directly in its cells instead of being hidden.
// Use Cases: Generate a financial report where division‑by‑zero or other formula errors must be visible for audit purposes. | Convert an existing workbook to expose data‑validation issues by toggling ShowErrorValues on all PivotTables. | Create automated Excel exports that retain error symbols (e.g., #DIV/0!) to help downstream users identify problematic rows.
// AI Prompts: Write C# code with Aspose.Cells that builds a PivotTable, adds a calculated field causing a #DIV/0! error, and sets ShowErrorValues to true. | Explain when to use DisplayErrorString versus ShowErrorValues in Aspose.Cells PivotTable settings. | Provide a step‑by‑step guide to enable error display for every PivotTable in an existing workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, build a PivotTable, define a calculated field that triggers a division‑by‑zero, and configure the PivotTable to show error values (using ShowErrorValues / DisplayErrorString) before refreshing and saving the file.
    public class PivotTableShowErrorValuesDemo
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
                cells["B2"].Value = 1200;
                cells["A3"].Value = "Travel";
                cells["B3"].Value = 800;
                cells["A4"].Value = "Food";
                cells["B4"].Value = 500;
                cells["A5"].Value = "Travel";
                cells["B5"].Value = 0; // Will cause division by zero later

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table (Category as row, Amount as data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount column

                // Introduce a calculated field that will generate an error (division by zero)
                pivotTable.AddCalculatedField("ErrorField", "='Amount'/0");

                // Configure the pivot table to display error strings
                pivotTable.DisplayErrorString = true;
                // Empty string shows the default Excel error (e.g., #DIV/0!)
                pivotTable.ErrorString = string.Empty;

                // Refresh the pivot table data to apply the calculated field
                pivotTable.RefreshData();

                // Recalculate the pivot table to reflect changes
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTableShowErrorValuesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}

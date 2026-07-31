// Title: Update a PivotTable calculated field with a discount rate using Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel workbook, finds the first PivotTable, retrieves an existing calculated field (e.g., TotalSales), creates a new formula that multiplies the original expression by a discount rate stored in cell D1, adds the new calculated field (TotalSalesWithDiscount), refreshes and recalculates the PivotTable, and saves the updated file.
// Keywords: Aspose.Cells | C# | PivotTable | calculated field | discount rate | update formula | add calculated field | RefreshData | CalculateData | Excel automation | Excel pivot table programming
// Common Searches: how to modify a calculated field formula in Aspose.Cells PivotTable | add discount rate to existing calculated field using Aspose.Cells .NET | create new calculated field based on another field in Aspose.Cells | refresh pivot table after adding calculated field Aspose.Cells | Aspose.Cells C# update pivot table formula with cell reference
// Developer Intent: Add a new calculated field that applies a discount rate to an existing PivotTable calculation without altering the original field.
// Use Cases: Generate a "TotalSalesWithDiscount" metric that applies the discount stored in D1 while preserving the original TotalSales field. | Introduce a dynamic discount factor for sales reporting that can be changed by editing a single worksheet cell. | Upgrade legacy reports by adding discounted sales calculations without breaking existing PivotTable layouts.
// AI Prompts: Write C# code with Aspose.Cells to update a PivotTable calculated field by incorporating a discount rate cell reference. | Explain step‑by‑step how to add a new calculated field, refresh the PivotTable, and save the workbook using Aspose.Cells. | Suggest best‑practice error handling for locating pivot tables and calculated fields before modifying formulas in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# example that loads an Excel workbook, finds the first PivotTable, retrieves an existing calculated field (e.g., TotalSales), creates a new formula that multiplies the original expression by a discount rate stored in cell D1, adds the new calculated field (TotalSalesWithDiscount), refreshes and recalculates the PivotTable, and saves the updated file.
    public class UpdateCalculatedFieldFormula
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure a pivot table exists
                if (sheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Assume the first pivot table is the target
                PivotTable pivotTable = sheet.PivotTables[0];

                string existingFieldName = "TotalSales";

                // Locate the existing calculated field in DataFields
                PivotField existingField = null;
                foreach (PivotField pf in pivotTable.DataFields)
                {
                    if (pf.Name == existingFieldName)
                    {
                        existingField = pf;
                        break;
                    }
                }

                if (existingField == null)
                {
                    Console.WriteLine($"Calculated field '{existingFieldName}' not found.");
                    return;
                }

                // Get current formula (e.g., "=Price*Quantity")
                string oldFormula = existingField.GetFormula();

                // Reference cell that holds the discount rate
                string discountRateReference = "D1";

                // Build new formula incorporating the discount rate
                string newFormula = $"=({oldFormula.TrimStart('=')})*{discountRateReference}";

                // Add a new calculated field with the updated formula
                string newFieldName = "TotalSalesWithDiscount";
                pivotTable.AddCalculatedField(newFieldName, newFormula, true);

                // Refresh and calculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}

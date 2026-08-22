// Title: Update an existing PivotTable calculated field formula to include a discount rate using Aspose.Cells for .NET (C#)
// AI Prompts: Locate a calculated field in a PivotTable with Aspose.Cells and replace its formula with an expression that references a DiscountRate source field. | Refresh the PivotTable, recalculate its data after the formula change, and save the workbook programmatically.
// Common Searches: Aspose.Cells C# change formula of a PivotTable calculated field to add discount rate | replace existing calculated field in Excel pivot using Aspose.Cells .NET | how to include a DiscountRate parameter in a PivotTable calculated field with Aspose.Cells | update pivot calculated field formula and refresh data programmatically in C# | Aspose.Cells add calculated field with same name to overwrite formula
// Tags: Aspose.Cells PivotTable calculated field update | C# modify pivot calculated field formula | add discount rate to pivot calculated field Aspose | refresh pivot after formula change Aspose.Cells | overwrite existing calculated field Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUpdateCalculatedField
{
    // The example loads an Excel workbook, accesses the first worksheet's PivotTable, finds a calculated field named "Total", builds a new formula that multiplies Price*Quantity by (1-DiscountRate), replaces the field by adding a calculated field with the same name, refreshes and recalculates the PivotTable, and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table with a calculated field
            Workbook workbook = new Workbook("input.xlsx"); // <-- replace with your source file
            Worksheet sheet = workbook.Worksheets[0];

            // Assume there is only one pivot table; get it
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found.");
                return;
            }

            PivotTable pivotTable = sheet.PivotTables[0];

            // Name of the calculated field we want to update
            string calcFieldName = "Total"; // <-- replace with your field name

            // Locate the calculated field in the DataFields collection
            PivotField calcField = null;
            foreach (PivotField field in pivotTable.DataFields)
            {
                if (field.IsCalculatedField && field.Name.Equals(calcFieldName, StringComparison.OrdinalIgnoreCase))
                {
                    calcField = field;
                    break;
                }
            }

            if (calcField == null)
            {
                Console.WriteLine($"Calculated field \"{calcFieldName}\" not found.");
                return;
            }

            // Retrieve the current formula (optional, for logging)
            string oldFormula = calcField.GetFormula();
            Console.WriteLine($"Old formula: {oldFormula}");

            // Define the new discount rate parameter name (must exist as a source field)
            string discountRateField = "DiscountRate";

            // Build the new formula incorporating the discount rate
            // Example: original formula was "=Price*Quantity"
            // New formula: "=Price*Quantity*(1-DiscountRate)"
            string newFormula = $"=Price*Quantity*(1-{discountRateField})";

            // Update the calculated field.
            // Aspose.Cells does not provide a direct SetFormula for PivotField,
            // but adding a calculated field with the same name replaces the existing one.
            pivotTable.AddCalculatedField(calcFieldName, newFormula, true);

            // Refresh and recalculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save("output.xlsx"); // <-- replace with your desired output file
            Console.WriteLine("Calculated field formula updated and workbook saved.");
        }
    }
}

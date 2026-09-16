// Title: How to reset an Excel PivotTable to the default (non‑compact) layout with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx workbook, finds the first PivotTable, and disables its compact layout by setting ShowInCompactForm to false, using reflection to handle different Aspose.Cells versions. | Show how to programmatically change a PivotTable's report layout to the standard form in Aspose.Cells, covering both property and method overloads for the ShowInCompactForm setting.
// Common Searches: Aspose.Cells C# set pivot table ShowInCompactForm false | disable compact layout for Excel pivot table using Aspose.Cells .NET | reset pivot table report layout to default with Aspose.Cells | reflection to change PivotTable layout in Aspose.Cells | how to change Excel pivot table layout programmatically in C#
// Tags: Aspose.Cells set ShowInCompactForm property | C# modify Excel pivot table layout | Aspose.Cells reflection for pivot table settings | disable compact form in Excel pivot table .NET | programmatic pivot table report layout Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads 'input.xlsx', checks for a PivotTable on the first worksheet, and uses reflection to set the PivotTable's ShowInCompactForm flag to false (or invoke the equivalent method), thereby switching the report layout to the default non‑compact form, and saves the modified workbook as 'output.xlsx'.
class PivotTableLayoutExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Set layout to non‑compact form.
            // Aspose.Cells may expose this as a property or a method depending on the version.
            // Use reflection to handle both possibilities safely.
            var ptType = typeof(PivotTable);
            var prop = ptType.GetProperty("ShowInCompactForm");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(pivotTable, false);
            }
            else
            {
                var method = ptType.GetMethod("ShowInCompactForm", new Type[] { typeof(bool) });
                method?.Invoke(pivotTable, new object[] { false });
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

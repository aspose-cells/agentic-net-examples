// Title: Disable PivotTable Column Auto‑Fit in Aspose.Cells for .NET (C#)
// Description: Loads a workbook, accesses the first worksheet’s pivot table, sets PivotTable.AutofitColumnWidthOnUpdate to false, refreshes the pivot tables to apply the change, and saves the file so custom column widths are retained.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | AutofitColumnWidthOnUpdate | disable auto fit | column width preservation | refresh pivot tables | Excel automation | programmatic pivot settings
// Common Searches: Aspose.Cells disable pivot column auto fit | C# set AutofitColumnWidthOnUpdate false | prevent pivot table column width change Aspose | keep custom column widths in pivot table using Aspose.Cells | refresh pivot tables after turning off autofit
// Developer Intent: Turn off automatic column‑width adjustment for a PivotTable to keep manually defined widths.
// Use Cases: Generate a report where column widths are predefined, then disable autofit before saving the workbook. | Apply the same setting to every PivotTable across multiple worksheets in a large Excel file. | Integrate the property change into an automated data‑export pipeline that creates pivot tables from raw data.
// AI Prompts: Write C# code with Aspose.Cells that disables column auto‑fit for all pivot tables in a workbook and refreshes them. | Explain the effect of the AutofitColumnWidthOnUpdate property and list other configurable PivotTable options in Aspose.Cells for .NET. | Provide step‑by‑step instructions to set custom column widths after turning off autofit for a pivot table.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace Example
{
    // Loads a workbook, accesses the first worksheet’s pivot table, sets PivotTable.AutofitColumnWidthOnUpdate to false, refreshes the pivot tables to apply the change, and saves the file so custom column widths are retained.
    class DisablePivotAutoFit
    {
        static void Main()
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

                // Ensure the workbook has at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Retrieve the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Disable automatic column width adjustment on update
                pivotTable.AutofitColumnWidthOnUpdate = false;

                // Refresh pivot tables to apply the setting
                worksheet.RefreshPivotTables();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

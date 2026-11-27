// Title: C# example: Delete Excel pivot tables whose names start with "Temp_" from every worksheet using Aspose.Cells
// Common Searches: C# Aspose.Cells remove pivot tables with name prefix Temp_ | How to delete temporary pivot tables from all sheets in an Excel file using Aspose.Cells | Iterate worksheets and delete pivot tables that begin with Temp_ in .NET | Aspose.Cells code to filter and remove pivot tables by naming convention | Delete pivot tables programmatically in C# based on name pattern
// Tags: remove pivot tables by prefix Aspose.Cells | delete temporary pivot tables C# | iterate worksheets to remove pivot tables .NET | pivot table deletion based on naming convention | Aspose.Cells filter pivot tables by name

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // PivotTable resides in this namespace in recent versions

// The sample loads an Excel workbook, loops through each worksheet, identifies pivot tables whose Name starts with "Temp_" (case‑insensitive), removes those pivot tables, and saves the workbook to a new file. It includes input‑file validation and exception handling for robust execution.
class DeleteTempPivotTables
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Collect pivot tables whose names start with "Temp_"
                List<PivotTable> tablesToRemove = new List<PivotTable>();
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    if (!string.IsNullOrEmpty(pt.Name) &&
                        pt.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                    {
                        tablesToRemove.Add(pt);
                    }
                }

                // Remove the collected pivot tables from the worksheet
                foreach (PivotTable pt in tablesToRemove)
                {
                    sheet.PivotTables.Remove(pt);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

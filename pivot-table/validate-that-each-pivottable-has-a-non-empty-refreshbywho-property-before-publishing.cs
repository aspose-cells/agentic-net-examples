// Title: Validate that every PivotTable has a non‑empty Name property before saving the workbook with Aspose.Cells for .NET
// Common Searches: how to ensure each PivotTable has a name using Aspose.Cells C# | validate PivotTable Name property before saving workbook in .NET | Aspose.Cells C# check for empty PivotTable names across worksheets | throw exception if PivotTable Name is missing with Aspose.Cells | iterate through worksheets to verify PivotTable names in C#
// Tags: Aspose.Cells PivotTable name validation | C# verify PivotTable properties | iterate worksheets Aspose.Cells | exception handling for empty PivotTable name | Excel workbook validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The example loads an Excel workbook with Aspose.Cells, iterates over each worksheet and its PivotTables, checks that the Name property of every PivotTable is not empty, throws an InvalidOperationException when a missing name is found, and saves the validated workbook to a new file.
class PivotTableValidator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pivot tables in the current worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Validate that the pivot table has a non‑empty name
                    if (string.IsNullOrEmpty(pivot.Name))
                    {
                        throw new InvalidOperationException(
                            $"A PivotTable in worksheet '{sheet.Name}' has an empty Name property.");
                    }
                }
            }

            // Save the validated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook validated and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Programmatically set the Outline layout for every PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to loop through all worksheets and assign the Outline layout to each PivotTable via reflection. | Write C# code that loads an .xlsx file, applies the Outline view to all pivot tables, and saves the updated workbook with Aspose.Cells. | Implement a bulk update that changes every PivotTable's LayoutType to Outline without referencing the enum directly, using reflection in Aspose.Cells.
// Common Searches: Aspose.Cells change pivot tables to outline view across all worksheets in C# | set PivotTable LayoutType using reflection with Aspose.Cells .NET | bulk update Excel pivot tables layout programmatically | iterate workbook and apply outline view to each pivot table C# | how to programmatically apply outline layout to every pivot table in an Excel file
// Tags: Aspose.Cells set PivotTable LayoutType outline | C# bulk modify Excel pivot tables layout | reflection update Aspose.Cells PivotTable properties | apply outline view to all PivotTables in workbook | iterate worksheets pivot tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel file, iterates through each worksheet and its PivotTables, uses reflection to set the LayoutType property to Outline for every pivot table, and saves the modified workbook, handling missing files and save errors.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets and their pivot tables
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Attempt to set layout to Outline using reflection (avoids compile‑time dependency)
                    try
                    {
                        var layoutProp = typeof(PivotTable).GetProperty("LayoutType");
                        // Get the enum type by its full name to avoid direct reference
                        var enumType = typeof(PivotTable).Assembly.GetType("Aspose.Cells.Pivot.PivotTableLayoutType");
                        if (layoutProp != null && enumType != null && layoutProp.CanWrite)
                        {
                            var enumValue = Enum.Parse(enumType, "Outline");
                            layoutProp.SetValue(pivot, enumValue);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unable to set layout for pivot table '{pivot.Name}': {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

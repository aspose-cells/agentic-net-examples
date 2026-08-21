// Title: Filter out defined names prefixed with "temp_" using Aspose.Cells LoadFilter (C#)
// Description: Demonstrates how to create a custom LoadFilter, load an Excel workbook with LoadOptions, and delete every defined name that begins with the "temp_" prefix before saving the file.
// Keywords: Aspose.Cells LoadFilter | LoadOptions C# | remove defined names | temporary named ranges | filter named ranges | temp_ prefix | Excel workbook cleanup | C# Aspose.Cells example
// Common Searches: Aspose.Cells delete defined names with prefix | C# load Excel file and skip temporary named ranges | How to use LoadFilter to filter named ranges in Aspose.Cells | Remove "temp_" named ranges from workbook using Aspose | LoadOptions LoadFilter example for named ranges
// Developer Intent: Remove all defined names that start with "temp_" from a workbook during or after loading.
// Use Cases: Strip placeholder named ranges before publishing a report. | Reduce file size by eliminating temporary names generated during data processing. | Prepare a clean workbook for downstream analytics that require only permanent named ranges.
// AI Prompts: Show C# code that uses Aspose.Cells LoadOptions with a custom LoadFilter to exclude defined names beginning with "temp_". | Explain how to iterate through Workbook.Worksheets.Names and delete entries that match a specific prefix. | Provide a step‑by‑step guide for cleaning up temporary named ranges in an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a custom LoadFilter, load an Excel workbook with LoadOptions, and delete every defined name that begins with the "temp_" prefix before saving the file.
class CustomLoadFilter : LoadFilter
{
    // Load all data for each worksheet.
    public override void StartSheet(Worksheet sheet)
    {
        LoadDataFilterOptions = LoadDataFilterOptions.All;
    }
}

class Program
{
    static void Main()
    {
        // Prepare load options and assign the custom filter.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new CustomLoadFilter();

        // Load the workbook using the specified options.
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Remove defined names that start with "temp_".
        for (int i = workbook.Worksheets.Names.Count - 1; i >= 0; i--)
        {
            Name definedName = workbook.Worksheets.Names[i];
            if (definedName.Text.StartsWith("temp_", StringComparison.OrdinalIgnoreCase))
            {
                workbook.Worksheets.Names.RemoveAt(i);
            }
        }

        // Save the filtered workbook.
        workbook.Save("output.xlsx");
    }
}

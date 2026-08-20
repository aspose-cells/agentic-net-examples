// Title: Load a workbook with Aspose.Cells .NET using LoadFilter to keep only specific defined names
// Description: Demonstrates how to create a LoadFilter that includes defined names and cell data, load a workbook, whitelist required named ranges (e.g., "SalesData" and "Expenses"), remove all other names, optionally sort the remaining names, and save the filtered file.
// Keywords: Aspose.Cells LoadFilter defined names | C# load specific named ranges | remove unwanted named ranges .NET | filter defined names Aspose.Cells | sort workbook names Aspose.Cells | memory optimization workbook loading | Aspose.Cells .NET example
// Common Searches: Aspose.Cells load only selected named ranges | How to filter defined names when opening a workbook in C# | Remove unused named ranges after loading workbook Aspose.Cells | Sort named ranges after filtering Aspose.Cells .NET | Reduce memory usage by loading specific defined names
// Developer Intent: Load a workbook and retain only the required named ranges while discarding all others.
// Use Cases: Load a template workbook and keep only the "SalesData" and "Expenses" ranges needed for calculations. | Minimize memory consumption by loading just the necessary defined names together with cell data. | Clean up a workbook before distribution by deleting unused named ranges and ordering the remaining ones.
// AI Prompts: Show me a C# example that uses Aspose.Cells LoadFilter to load only certain defined names and delete the rest. | Provide code to whitelist specific named ranges, remove all others, sort the names, and save the workbook with Aspose.Cells .NET. | Explain how combining LoadDataFilterOptions.DefinedNames and LoadDataFilterOptions.CellData improves performance when opening large workbooks.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to create a LoadFilter that includes defined names and cell data, load a workbook, whitelist required named ranges (e.g., "SalesData" and "Expenses"), remove all other names, optionally sort the remaining names, and save the filtered file.
class Program
{
    static void Main()
    {
        // Define the names of the ranges that must be kept after loading
        var requiredNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SalesData",
            "Expenses"
        };

        // Create a LoadFilter that loads defined names (and cell data if needed)
        LoadFilter loadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames | LoadDataFilterOptions.CellData);

        // Assign the filter to LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = loadFilter;

        // Load the workbook using the specified options
        Workbook workbook = new Workbook("Template.xlsx", loadOptions);

        // Retrieve all defined names (workbook‑scoped and worksheet‑scoped)
        Name[] allNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

        // Remove any name that is not in the required list
        foreach (Name name in allNames)
        {
            if (!requiredNames.Contains(name.Text))
            {
                workbook.Worksheets.Names.Remove(name.Text);
            }
        }

        // Optional: sort the remaining names for consistency
        workbook.Worksheets.SortNames();

        // Save the filtered workbook
        workbook.Save("FilteredWorkbook.xlsx");
    }
}

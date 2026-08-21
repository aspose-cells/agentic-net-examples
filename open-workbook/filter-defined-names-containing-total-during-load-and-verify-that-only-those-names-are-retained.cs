// Title: C# Aspose.Cells: Load Workbook and Keep Only Defined Names Containing “Total”
// Description: Shows how to use Aspose.Cells LoadOptions with a custom LoadFilter to load only defined names, delete every named range that lacks the word “Total”, confirm the remaining names, and save the filtered workbook.
// Keywords: Aspose.Cells | C# LoadFilter | defined names | named ranges | filter by keyword | Total named ranges | load workbook | remove named ranges | Excel automation | memory optimization
// Common Searches: Aspose.Cells load only defined names C# | remove named ranges except those containing a word | filter Excel named ranges by keyword Aspose | verify remaining defined names after LoadFilter | C# load workbook with custom LoadFilter
// Developer Intent: Load an Excel file, retain only defined names that include the word “Total”, and save the trimmed workbook.
// Use Cases: Process large financial workbooks while loading only total‑related named ranges to lower memory consumption. | Create a lightweight copy of a template that contains just summary named ranges for downstream reporting. | Validate that a workbook includes required total named ranges before running calculations.
// AI Prompts: Write a C# LoadFilter subclass for Aspose.Cells that loads only defined names during workbook loading. | Generate C# code that removes all defined names not containing a specified substring after loading an Excel file with Aspose.Cells. | Create a unit test that asserts only names containing "Total" remain after applying the custom LoadFilter and removal logic.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Shows how to use Aspose.Cells LoadOptions with a custom LoadFilter to load only defined names, delete every named range that lacks the word “Total”, confirm the remaining names, and save the filtered workbook.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputFile = "source.xlsx";
        string outputFile = "filtered.xlsx";

        // Create LoadOptions and assign a custom LoadFilter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new TotalNamesLoadFilter();

        // Load the workbook using the specified LoadOptions
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // After loading, remove all defined names that do NOT contain "Total"
        NameCollection names = workbook.Worksheets.Names;
        List<string> namesToRemove = new List<string>();

        foreach (Name name in names)
        {
            if (!name.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
            {
                namesToRemove.Add(name.Text);
            }
        }

        if (namesToRemove.Count > 0)
        {
            names.Remove(namesToRemove.ToArray());
        }

        // Verify that only names containing "Total" remain
        Console.WriteLine("Remaining defined names after filtering:");
        foreach (Name name in names)
        {
            Console.WriteLine(name.Text);
        }

        // Save the filtered workbook
        workbook.Save(outputFile);
    }

    // Custom LoadFilter that loads only defined names during workbook loading
    class TotalNamesLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only the defined names (other data can be loaded as needed)
            LoadDataFilterOptions = LoadDataFilterOptions.DefinedNames;
        }
    }
}

// Title: C# – Load an Excel file with Aspose.Cells using LoadFilter to read only defined names
// Description: Shows how to open an .xlsx workbook with Aspose.Cells, apply a LoadFilter (LoadDataFilterOptions.DefinedNames) via LoadOptions, and retrieve every workbook‑ and worksheet‑scoped name without loading any cell data. The sample prints each name and its reference.
// Keywords: Aspose.Cells LoadFilter | LoadDataFilterOptions.DefinedNames | C# read defined names only | Excel named ranges without worksheets | Workbook.Names.Filter | Aspose.Cells performance | skip worksheet data | named ranges extraction | Aspose.Cells GitHub example | C# Excel API
// Common Searches: Aspose.Cells load only defined names | C# read named ranges without loading sheets | How to use LoadFilter for defined names in Aspose.Cells | Retrieve workbook scoped names Aspose.Cells | Get Excel named ranges without cell data C# | Performance tip Aspose.Cells load filter defined names | GitHub Aspose.Cells defined names sample
// Developer Intent: The developer wants to open an Excel workbook and obtain all defined names while avoiding the overhead of loading worksheet cell data.
// Use Cases: Extract named ranges for validation or reporting without processing large sheet contents. | Generate documentation or mapping tables that list every defined name in a workbook. | Quickly verify the existence of specific named ranges in a massive file without incurring worksheet load time.
// AI Prompts: Provide a C# example that loads only defined names from an Excel file using Aspose.Cells and returns them as a collection. | Explain how to adjust the LoadFilter to include both defined names and workbook properties while still skipping worksheet cell data. | Show how to filter only workbook‑scoped names after loading with a DefinedNames filter using the Names.Filter method.

using System;
using Aspose.Cells;

// Shows how to open an .xlsx workbook with Aspose.Cells, apply a LoadFilter (LoadDataFilterOptions.DefinedNames) via LoadOptions, and retrieve every workbook‑ and worksheet‑scoped name without loading any cell data. The sample prints each name and its reference.
class Program
{
    static void Main()
    {
        // Path to the workbook file to be loaded
        string filePath = "input.xlsx";

        // Create a LoadFilter that loads only defined name objects
        LoadFilter filter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

        // Assign the filter to LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = filter;

        // Load the workbook using the specified LoadOptions
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Retrieve all defined names (workbook‑scoped, worksheet‑scoped, etc.)
        Name[] definedNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

        // Output the retrieved names
        Console.WriteLine($"Total defined names: {definedNames.Length}");
        foreach (Name name in definedNames)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }
    }
}

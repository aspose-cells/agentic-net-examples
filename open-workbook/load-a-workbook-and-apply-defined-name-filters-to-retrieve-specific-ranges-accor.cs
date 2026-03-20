using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class DefinedNameFilterDemo
{
    static void Main()
    {
        // Load a workbook, but load only defined name objects to reduce memory usage
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);
        Workbook workbook = new Workbook("Input.xlsx", loadOptions);

        // ------------------------------------------------------------
        // 1. Retrieve defined names by different scopes
        // ------------------------------------------------------------
        // Workbook‑scoped names (global names)
        Name[] workbookScopeNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);
        Console.WriteLine($"Workbook‑scoped names count: {workbookScopeNames.Length}");

        // Worksheet‑scoped names for all worksheets
        Name[] worksheetScopeAll = workbook.Worksheets.Names.Filter(NameScopeType.Worksheet, -1);
        Console.WriteLine($"Worksheet‑scoped names (all sheets) count: {worksheetScopeAll.Length}");

        // Worksheet‑scoped names for a specific sheet (index 0)
        Name[] sheet0Names = workbook.Worksheets.Names.Filter(NameScopeType.Worksheet, 0);
        Console.WriteLine($"Worksheet‑scoped names for sheet index 0 count: {sheet0Names.Length}");

        // ------------------------------------------------------------
        // 2. Work with each defined name – obtain its range and display info
        // ------------------------------------------------------------
        foreach (Name definedName in workbookScopeNames)
        {
            // Get the range the name refers to
            AsposeRange range = definedName.GetRange();

            if (range != null)
            {
                Console.WriteLine($"Name '{definedName.Text}' refers to range {range.RefersTo} on sheet '{range.Worksheet.Name}'.");
            }
            else
            {
                Console.WriteLine($"Name '{definedName.Text}' does not refer to a valid range.");
            }
        }

        // ------------------------------------------------------------
        // 3. Retrieve a range directly by its name using WorksheetCollection
        // ------------------------------------------------------------
        if (workbookScopeNames.Length > 0)
        {
            string firstName = workbookScopeNames[0].Text;
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName(firstName);

            if (namedRange != null)
            {
                Console.WriteLine($"GetRangeByName(\"{firstName}\") returned address: {namedRange.Address}");
            }
            else
            {
                Console.WriteLine($"GetRangeByName could not find a range for name \"{firstName}\".");
            }
        }

        // ------------------------------------------------------------
        // 4. Save the workbook (no modifications made, just to demonstrate saving)
        // ------------------------------------------------------------
        workbook.Save("Output.xlsx");
    }
}
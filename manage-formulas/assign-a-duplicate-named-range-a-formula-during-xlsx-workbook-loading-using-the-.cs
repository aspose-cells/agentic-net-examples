using System;
using System.Collections.Generic;
using Aspose.Cells;

class AssignFormulaToDuplicateNamedRanges
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to track first occurrence of each name text
        var nameTracker = new Dictionary<string, Name>(StringComparer.OrdinalIgnoreCase);

        // Iterate through all defined names in the workbook
        foreach (Name name in workbook.Worksheets.Names)
        {
            // If the name has not been seen before, store it
            if (!nameTracker.ContainsKey(name.Text))
            {
                nameTracker[name.Text] = name;
                continue;
            }

            // Duplicate detected – assign a formula to this duplicate name
            // Example formula: refer to cell A1 on the first worksheet
            // Ensure the formula starts with an equal sign as required by RefersTo
            name.RefersTo = $"=Sheet1!$A$1";
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
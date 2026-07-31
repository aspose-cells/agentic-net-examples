// Title: Check for duplicate named ranges in an Aspose.Cells workbook (C#)
// Description: This C# example creates or loads a workbook, adds named ranges (including intentional duplicates), scans the workbook's NameCollection case‑insensitively, reports any repeated names to the console, and saves the file.
// Keywords: Aspose.Cells | C# | named range validation | duplicate detection | NameCollection | Excel workbook | unique range names | range conflict | programmatic Excel | cell naming
// Common Searches: Aspose.Cells find duplicate names | C# detect repeated named ranges | validate named range uniqueness Excel | how to list duplicate names Aspose | check NameCollection for duplicates
// Developer Intent: Identify and list any named ranges that share the same identifier within a workbook.
// Use Cases: Run an automated quality‑check before distributing generated reports to avoid reference errors. | Prevent runtime failures when formulas point to ambiguous named ranges. | Enforce naming conventions in large financial or analytical models.
// AI Prompts: Generate a method that returns a List<string> of duplicate named range identifiers from a Workbook using Aspose.Cells. | Refactor the sample to throw a DuplicateNameException instead of writing duplicate information to the console. | Write NUnit unit tests that verify case‑insensitive duplicate detection works correctly.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// This C# example creates or loads a workbook, adds named ranges (including intentional duplicates), scans the workbook's NameCollection case‑insensitively, reports any repeated names to the console, and saves the file.
class ValidateNamedRanges
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Sample data: add some named ranges, including duplicates
        // -------------------------------------------------
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Value1");
        ws.Cells["B1"].PutValue("Value2");

        // First occurrence of "DupName"
        int idx1 = workbook.Worksheets.Names.Add("DupName");
        workbook.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1";

        // A unique name
        int idx2 = workbook.Worksheets.Names.Add("UniqueName");
        workbook.Worksheets.Names[idx2].RefersTo = "=Sheet1!$B$1";

        // Second occurrence of "DupName" (duplicate)
        int idx3 = workbook.Worksheets.Names.Add("DupName");
        workbook.Worksheets.Names[idx3].RefersTo = "=Sheet1!$A$2";

        // -------------------------------------------------
        // Validation: check for duplicate named range names
        // -------------------------------------------------
        NameCollection nameColl = workbook.Worksheets.Names;
        var nameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (Name name in nameColl)
        {
            string txt = name.Text;
            if (nameCounts.ContainsKey(txt))
                nameCounts[txt]++;
            else
                nameCounts[txt] = 1;
        }

        bool hasDuplicates = false;
        Console.WriteLine("Duplicate named ranges:");
        foreach (var kvp in nameCounts)
        {
            if (kvp.Value > 1)
            {
                hasDuplicates = true;
                Console.WriteLine($"- {kvp.Key} appears {kvp.Value} times");
            }
        }

        if (!hasDuplicates)
        {
            Console.WriteLine("No duplicate named ranges found.");
        }

        // -------------------------------------------------
        // Save the workbook (optional: after validation)
        // -------------------------------------------------
        workbook.Save("ValidatedWorkbook.xlsx");
    }
}

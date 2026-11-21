// Title: Detect and repair corrupted named ranges caused by worksheet renaming with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells code to delete named ranges that reference missing worksheets | how to fix broken named range RefersTo after renaming a sheet in Excel using Aspose.Cells | Aspose.Cells .NET example for validating and cleaning up corrupted named ranges | remove invalid named ranges from an Excel workbook with Aspose.Cells C# | detect missing sheet references in named ranges using the Aspose.Cells API
// Tags: Aspose.Cells named range corruption detection | C# remove invalid named ranges | validate RefersTo sheet references Aspose.Cells | Excel worksheet rename handling named ranges | NameCollection cleanup Aspose.Cells | repair broken named range references C#

using System;
using Aspose.Cells;

// The sample loads an Excel workbook, iterates through its NameCollection, parses each named range's RefersTo formula to extract the sheet name, checks if that sheet exists, and removes the named range when the referenced sheet is missing, then saves the corrected workbook.
class NamedRangeRepair
{
    static void Main()
    {
        // Load the workbook (replace with actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of named ranges in the workbook
        NameCollection names = workbook.Worksheets.Names;

        // Iterate over a copy of the names list because we may modify the collection
        for (int i = names.Count - 1; i >= 0; i--)
        {
            Name name = names[i];
            string refersTo = name.RefersTo; // e.g. =Sheet1!$A$1:$B$2

            // Skip if the reference is empty or does not contain a sheet delimiter
            if (string.IsNullOrEmpty(refersTo) || !refersTo.Contains("!"))
                continue;

            // Extract the sheet name part (remove leading '=' and any surrounding quotes)
            string sheetPart = refersTo.Substring(1, refersTo.IndexOf('!') - 1).Trim('\'', '\"');

            // Check if the sheet exists in the workbook
            bool sheetExists = false;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (string.Equals(ws.Name, sheetPart, StringComparison.OrdinalIgnoreCase))
                {
                    sheetExists = true;
                    break;
                }
            }

            // If the referenced sheet does not exist, the named range is corrupted
            if (!sheetExists)
            {
                // Option 1: Remove the corrupted named range
                names.RemoveAt(i);
                // Optionally, you could log or handle it differently here
                continue;
            }

            // If the sheet exists but the reference uses an old sheet name (e.g., renamed),
            // you can update the reference to the current sheet name.
            // This example assumes the sheet name in the reference is correct,
            // so no further action is taken.
        }

        // Save the corrected workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}

// Title: Renumber worksheet TabId values sequentially and ensure uniqueness with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that loads an Excel workbook, throws an exception if any worksheet has a duplicate TabId, then reassigns TabId values from 0 upward and saves the result. | Create a method that iterates through all worksheets, validates that each TabId is unique, updates each TabId to a sequential index, and returns the modified workbook.
// Common Searches: C# Aspose.Cells how to check for duplicate worksheet TabId | Aspose.Cells assign incremental TabId to each sheet | ensure unique TabId values before saving workbook Aspose.Cells .NET | renumber Excel sheet tabs programmatically using Aspose.Cells | validate worksheet TabId uniqueness in C#
// Tags: worksheet TabId sequential assignment Aspose.Cells | duplicate TabId detection Aspose.Cells | C# validate unique sheet TabId | Aspose.Cells set TabId property | renumber Excel sheet tabs .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Loads an Excel workbook, verifies that each worksheet's TabId is unique, reassigns TabIds sequentially starting at 0, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook
        var workbook = new Workbook("input.xlsx");

        // Access all worksheets
        var worksheets = workbook.Worksheets;

        // Verify that existing TabIds are unique
        var existingIds = new HashSet<int>();
        foreach (Worksheet sheet in worksheets)
        {
            int id = sheet.TabId;
            if (!existingIds.Add(id))
            {
                throw new InvalidOperationException($"Duplicate TabId {id} found in worksheet \"{sheet.Name}\".");
            }
        }

        // Assign sequential TabIds (starting from 0)
        for (int i = 0; i < worksheets.Count; i++)
        {
            worksheets[i].TabId = i;
        }

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}

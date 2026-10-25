// Title: Assign sequential TabId values to all worksheets, ensure uniqueness, and save the workbook using Aspose.Cells for .NET
// Common Searches: C# assign incremental TabId to each worksheet with Aspose.Cells | how to check for duplicate worksheet TabId in an Excel file using Aspose.Cells | update worksheet TabId property sequentially and save workbook in .NET | Aspose.Cells verify unique TabId across multiple sheets | detect duplicate TabId values in Excel workbook programmatically C#
// Tags: sequential TabId assignment for worksheets Aspose.Cells | worksheet TabId uniqueness check C# | modify worksheet TabId property Aspose.Cells | save workbook after TabId update Aspose.Cells | identify repeated TabId values across worksheets C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook, scans all worksheets to report any duplicate TabId values, reassigns TabId properties sequentially starting at 1, and then saves the updated workbook to a new file.
class UpdateTabIds
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Collect existing TabIds to verify uniqueness
        HashSet<int> existingIds = new HashSet<int>();
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Assuming Worksheet has a TabId property (int)
            int tabId = sheet.TabId;

            // If a duplicate is found, report it
            if (!existingIds.Add(tabId))
            {
                Console.WriteLine($"Duplicate TabId detected: {tabId} on sheet \"{sheet.Name}\"");
            }
        }

        // Assign sequential TabIds starting from 1
        int sequentialId = 1;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.TabId = sequentialId;
            sequentialId++;
        }

        // Save the updated workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}

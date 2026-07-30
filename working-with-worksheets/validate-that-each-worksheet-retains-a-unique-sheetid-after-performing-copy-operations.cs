// Title: Ensure Unique TabId and UniqueId for Worksheets After AddCopy in Aspose.Cells for .NET
// Description: C# sample that creates a workbook, adds two original sheets, copies them with Worksheets.AddCopy (by name and by index), renames the copies, and validates that every worksheet retains a distinct TabId and UniqueId using HashSet collections before saving the file.
// Keywords: Aspose.Cells | C# | .NET | worksheet TabId | UniqueId | AddCopy | copy worksheet | duplicate sheet ID | Excel automation | sample code
// Common Searches: Aspose.Cells verify unique TabId after AddCopy | detect duplicate UniqueId in copied worksheets .NET | how to ensure worksheet IDs are unique in Aspose.Cells | C# example for checking sheet identifiers after copy | Aspose.Cells AddCopy duplicate ID issue
// Developer Intent: Confirm that each worksheet in a workbook has a separate TabId and UniqueId after performing copy operations.
// Use Cases: Run a post‑copy integrity check to prevent ID collisions in merged workbooks. | Log or raise an alert when a duplicate TabId or UniqueId is detected during automated report generation. | Integrate the validation step into CI pipelines that generate Excel files with dynamic sheet duplication.
// AI Prompts: Generate a C# function that returns true only if all worksheets in an Aspose.Cells workbook have unique TabId and UniqueId values. | Write code that throws a custom exception when a duplicate TabId is found after using Worksheets.AddCopy. | Create an xUnit test verifying that copying worksheets with AddCopy never produces duplicate identifiers.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// C# sample that creates a workbook, adds two original sheets, copies them with Worksheets.AddCopy (by name and by index), renames the copies, and validates that every worksheet retains a distinct TabId and UniqueId using HashSet collections before saving the file.
class ValidateSheetIds
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add initial worksheets
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Original1";
        sheet1.Cells["A1"].PutValue("Data1");

        Worksheet sheet2 = workbook.Worksheets.Add("Original2");
        sheet2.Cells["A1"].PutValue("Data2");

        // Copy worksheets using AddCopy (by name and by index)
        int copyIndex1 = workbook.Worksheets.AddCopy("Original1");
        Worksheet copy1 = workbook.Worksheets[copyIndex1];
        copy1.Name = "Copy1";

        int copyIndex2 = workbook.Worksheets.AddCopy(1); // copy "Original2" by index
        Worksheet copy2 = workbook.Worksheets[copyIndex2];
        copy2.Name = "Copy2";

        // Validate that each worksheet has a unique TabId
        HashSet<int> tabIds = new HashSet<int>();
        bool allUnique = true;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            int id = ws.TabId;
            if (!tabIds.Add(id))
            {
                Console.WriteLine($"Duplicate TabId found: {id} on worksheet '{ws.Name}'");
                allUnique = false;
            }
        }

        if (allUnique)
        {
            Console.WriteLine("All worksheets have unique TabId values.");
        }

        // Additionally, validate that each worksheet has a unique UniqueId
        HashSet<string> uniqueIds = new HashSet<string>();
        foreach (Worksheet ws in workbook.Worksheets)
        {
            string uid = ws.UniqueId;
            if (!uniqueIds.Add(uid))
            {
                Console.WriteLine($"Duplicate UniqueId found: {uid} on worksheet '{ws.Name}'");
                allUnique = false;
            }
        }

        if (allUnique)
        {
            Console.WriteLine("All worksheets have unique UniqueId values.");
        }

        // Save the workbook (optional)
        workbook.Save("ValidateSheetIds.xlsx");
    }
}

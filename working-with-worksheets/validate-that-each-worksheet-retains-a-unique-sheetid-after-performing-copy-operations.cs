// Title: Verify that each worksheet retains a distinct Index after copying sheets with Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies several worksheets using Aspose.Cells AddCopy and then iterates the Worksheets collection to confirm every worksheet's Index is unique. | Write a console application that adds new sheets, creates copies with AddCopy, and reports any duplicate Index values in an Aspose.Cells workbook. | Create a method that validates worksheet identifiers after copy operations by storing each worksheet's Index in a HashSet and flagging repeats.
// Common Searches: asp.net how to verify worksheet index uniqueness after using AddCopy in Aspose.Cells | c# detect duplicate sheet indexes in a workbook created with Aspose.Cells | ensure copied worksheets have different IDs in Aspose.Cells .NET | Aspose.Cells AddCopy duplicate Index issue solution | programmatically check for repeated worksheet indices after copying sheets in C#
// Tags: Aspose.Cells AddCopy unique Index validation | C# worksheet duplicate index detection | hashset sheet identifier check Aspose.Cells | validate worksheet IDs after copy operation | Aspose.Cells workbook sheet identifier uniqueness

using System;
using System.Collections.Generic;
using Aspose.Cells;

// A C# console program that creates a workbook, adds and copies worksheets with AddCopy, then uses a HashSet to ensure each worksheet's Index (used as SheetId) remains unique, outputting any duplicates found.
class SheetIdValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();

            // Rename the default worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "OriginalSheet";

            // Add additional worksheets for testing
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");

            // Copy the first sheet and rename the copy
            int copyIndex1 = workbook.Worksheets.AddCopy(sheet1.Index);
            Worksheet copyOfOriginal = workbook.Worksheets[copyIndex1];
            copyOfOriginal.Name = "CopyOfOriginal";

            // Copy the second sheet and rename the copy
            int copyIndex2 = workbook.Worksheets.AddCopy(sheet2.Index);
            Worksheet copyOfSecond = workbook.Worksheets[copyIndex2];
            copyOfSecond.Name = "CopyOfSecond";

            // Validate that each worksheet retains a unique identifier (using Index)
            HashSet<int> sheetIds = new HashSet<int>();
            bool duplicateFound = false;

            foreach (Worksheet ws in workbook.Worksheets)
            {
                int id = ws.Index; // Index serves as a unique identifier
                if (!sheetIds.Add(id))
                {
                    duplicateFound = true;
                    Console.WriteLine($"Duplicate Index found: {id} in worksheet \"{ws.Name}\"");
                }
            }

            if (!duplicateFound)
            {
                Console.WriteLine("All worksheets have unique Index values after copy operations.");
            }

            // Optional: Save the workbook if needed
            // workbook.Save("CopyValidationResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

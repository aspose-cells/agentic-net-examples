// Title: How to protect an Excel worksheet while allowing edits only in a named range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect worksheet while keeping a specific named range editable | how to unlock cells in a named range before protecting a sheet with Aspose.Cells | C# example to protect Excel sheet and allow edits only in the 'EditableSection' range using Aspose.Cells | retrieve a named range by name and set IsLocked false with Aspose.Cells .NET | apply an unlocked style to a range and then protect the worksheet using ProtectionType.All in Aspose.Cells
// Tags: worksheet protection with selective unlocked range Aspose.Cells | style.IsLocked false for named range Aspose.Cells | range.ApplyStyle using StyleFlag.Locked Aspose.Cells | Workbook.GetRangeByName to retrieve named range Aspose.Cells | sheet.Protect with ProtectionType.All Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an existing workbook, retrieves the named range "EditableSection", creates a style with IsLocked set to false, applies this unlocked style to the range using a StyleFlag, then protects the worksheet with ProtectionType.All so only the unlocked cells remain editable, and finally saves the protected workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range "EditableSection"
            Aspose.Cells.Range editableRange = workbook.Worksheets.GetRangeByName("EditableSection");

            if (editableRange != null)
            {
                // Create a style where cells are unlocked
                Style unlockedStyle = workbook.CreateStyle();
                unlockedStyle.IsLocked = false; // Unlock cells

                // Specify that only the Locked attribute should be applied
                StyleFlag flag = new StyleFlag();
                flag.Locked = true;

                // Apply the unlocked style to all cells in the named range
                editableRange.ApplyStyle(unlockedStyle, flag);
            }
            else
            {
                Console.WriteLine("Named range 'EditableSection' not found.");
            }

            // Protect the worksheet; only unlocked cells can be edited
            sheet.Protect(ProtectionType.All);

            // Save the workbook with the protection applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

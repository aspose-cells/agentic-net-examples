// Title: How to protect an Excel worksheet with Aspose.Cells for .NET while keeping only the A1:B2 range editable
// AI Prompts: Write C# code using Aspose.Cells to lock an entire worksheet with a password but unlock cells A1:B2 for editing. | Show the steps to create an unlocked style, apply it to a specific range, and then call Worksheet.Protect with ProtectionType.All in Aspose.Cells. | Generate a complete example that creates a workbook, defines a named editable range, unlocks it, protects the sheet, and saves the file.
// Common Searches: Aspose.Cells C# protect worksheet and allow editing only in a specific range | How to unlock cells A1:B2 before applying worksheet protection with Aspose.Cells | C# example for setting password protection on Excel sheet while keeping certain cells editable using Aspose.Cells | Apply unlocked style to a range then protect sheet Aspose.Cells .NET | Worksheet.Protect with ProtectionType.All and unlocked cells Aspose.Cells tutorial
// Tags: worksheet protection with unlocked range Aspose.Cells | unlock cells A1:B2 using StyleFlag Aspose.Cells | apply password to Excel workbook Aspose.Cells .NET | create unlocked style for specific range Aspose.Cells | protect sheet while allowing edits in defined range C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, unlocks cells A1:B2 by applying an unlocked style, then protects the entire worksheet with a password using Worksheet.Protect, allowing only the specified range to be edited, and saves the file as an .xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range A1:B2 that should remain editable (rows 0-1, columns 0-1)
            Aspose.Cells.Range editableRange = sheet.Cells.CreateRange(0, 0, 2, 2);

            // Create a style that unlocks cells
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;

            // Apply the style to the defined range (apply all style attributes)
            StyleFlag flag = new StyleFlag { All = true };
            editableRange.ApplyStyle(unlockedStyle, flag);

            // Protect the worksheet while allowing edits only in unlocked cells
            // The Protect method requires the old password parameter; pass null if not needed
            sheet.Protect(ProtectionType.All, "YourPassword", null);

            // Save the protected workbook
            string outputPath = "ProtectedWorkbook.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

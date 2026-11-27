// Title: How to unprotect an Excel workbook structure, rename a worksheet, and protect it again with a new password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unprotect workbook structure with existing password then rename sheet | rename first worksheet after removing protection using Aspose.Cells .NET | apply new structure protection password to Excel file with Aspose.Cells | C# code to change worksheet name and re‑protect workbook structure in Aspose.Cells | update workbook protection password after modifying sheet name using Aspose.Cells
// Tags: unprotect workbook structure Aspose.Cells .NET | rename worksheet programmatically Aspose.Cells | protect workbook structure with new password Aspose.Cells | Aspose.Cells workbook protection and rename workflow | C# change sheet name after unprotecting workbook

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing Excel file, removes its structure protection using the old password, renames the first worksheet, reapplies structure protection with a new password, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook using the existing password
            string oldPassword = "oldPass";
            workbook.Unprotect(oldPassword);

            // Rename the first worksheet (index 0) to a new name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "RenamedSheet";

            // Re‑apply structure protection with a new password
            string newPassword = "newPass";
            workbook.Protect(ProtectionType.Structure, newPassword);

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

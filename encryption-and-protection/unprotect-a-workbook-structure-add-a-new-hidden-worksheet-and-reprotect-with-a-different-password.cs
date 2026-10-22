// Title: How to unprotect an Excel workbook structure, insert a hidden worksheet, and re‑protect it with a new password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove workbook structure protection and add a hidden sheet | C# code to unprotect an Excel file, hide a newly added worksheet, then protect the structure with a new password | programmatically change workbook protection password after inserting a hidden worksheet using Aspose.Cells | load existing XLSX, unprotect, add hidden worksheet, and protect structure with Aspose.Cells .NET | using Workbook.Unprotect and Workbook.Protect to modify Excel protection in C#
// Tags: Workbook.Unprotect method Aspose.Cells | hidden worksheet insertion Aspose.Cells | Workbook.Protect structure password .NET | Excel workbook protection workflow Aspose.Cells | change workbook protection password C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file, removes its workbook‑structure protection with the old password, adds a new hidden worksheet, then re‑applies structure protection using a new password, and saves the updated workbook.
class WorkbookProtectionExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string oldPassword = "oldPassword";   // replace with actual password if needed
        const string newPassword = "newPassword";   // replace with desired password

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook (removes all protection, including structure)
            // If the workbook is not password‑protected, pass null or an empty string
            workbook.Unprotect(string.IsNullOrEmpty(oldPassword) ? null : oldPassword);

            // Add a new worksheet
            int newSheetIndex = workbook.Worksheets.Add();

            // Get reference to the newly added worksheet and hide it
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.IsVisible = false;

            // Protect only the workbook structure with a new password
            workbook.Protect(ProtectionType.Structure, newPassword);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: How to unprotect an Excel workbook structure, add a hidden worksheet, and re‑protect it with a new password using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove workbook structure protection with old password | Add a hidden sheet to an existing Excel file using Aspose.Cells .NET | Change workbook protection password after inserting a hidden worksheet with Aspose.Cells | Aspose.Cells protect workbook structure with a new password after modifying worksheets | Unprotect and re‑protect Excel workbook structure programmatically in C#
// Tags: unprotect workbook structure Aspose.Cells | add hidden worksheet Aspose.Cells C# | protect workbook structure with password Aspose.Cells | modify Excel file hidden sheet Aspose.Cells | change workbook protection password .NET

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing Excel workbook, removes its structure protection using the current password, adds a new hidden worksheet named "HiddenSheet", re‑applies structure protection with a new password, and saves the updated file.
class WorkbookProtectionExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string oldPassword = "oldPassword"; // replace with actual password
        const string newPassword = "newPassword"; // replace with desired password

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Unprotect the workbook structure using the current password
            workbook.Unprotect(oldPassword);

            // Add a new worksheet and set its name
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "HiddenSheet";

            // Hide the newly added worksheet (using IsVisible for compatibility)
            newSheet.IsVisible = false;

            // Re‑protect the workbook structure with a new password
            workbook.Protect(ProtectionType.Structure, newPassword);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

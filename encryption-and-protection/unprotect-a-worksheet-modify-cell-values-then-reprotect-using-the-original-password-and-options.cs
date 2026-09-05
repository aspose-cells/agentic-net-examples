// Title: Unprotect an Excel worksheet, edit specific cells, and re‑protect it with the original password and custom protection settings using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx workbook, calls Worksheet.Unprotect with a password, updates cell values, configures desired Protection flags, and then re‑applies Worksheet.Protect with the same password. | Generate a C# example that demonstrates removing worksheet protection, changing cells A1 and B2, setting AllowFiltering, AllowFormattingCell, and other Allow* options, and saving the modified workbook.
// Common Searches: Aspose.Cells C# unprotect worksheet edit cells protect again with same password | C# code to change values in a protected Excel sheet using Aspose.Cells | How to keep custom protection options when re‑protecting a worksheet in Aspose.Cells | Set specific Allow* protection flags on an Excel worksheet with Aspose.Cells .NET | Programmatically unprotect and protect Excel worksheet with password using Aspose.Cells
// Tags: worksheet unprotect Aspose.Cells C# | modify cell values protected worksheet Aspose.Cells | worksheet protect custom options Aspose.Cells | Aspose.Cells protection flags configuration | C# edit protected Excel file Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads input.xlsx, unprotects the first worksheet with a password, updates cells A1 and B2, configures various Protection flags (e.g., AllowFiltering, AllowFormattingCell), re‑protects the sheet using the same password, and saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "myPassword";

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

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unprotect the worksheet using the original password
            sheet.Unprotect(password);

            // ----- Modify cell values -----
            sheet.Cells["A1"].PutValue("Updated Text");
            sheet.Cells["B2"].PutValue(12345);
            // Add more modifications as needed

            // ----- Re‑protect the worksheet with the same password and options -----
            // Configure protection options directly via the Protection property
            sheet.Protection.AllowDeletingColumn = false;
            sheet.Protection.AllowDeletingRow = false;
            sheet.Protection.AllowEditingObject = false;
            sheet.Protection.AllowEditingScenario = false;
            sheet.Protection.AllowFiltering = true;
            sheet.Protection.AllowFormattingCell = true;
            sheet.Protection.AllowFormattingColumn = true;
            sheet.Protection.AllowFormattingRow = true;
            sheet.Protection.AllowInsertingColumn = false;
            sheet.Protection.AllowInsertingHyperlink = false;
            sheet.Protection.AllowInsertingRow = false;
            sheet.Protection.AllowSelectingLockedCell = true;
            sheet.Protection.AllowSelectingUnlockedCell = true;
            sheet.Protection.AllowSorting = true;

            // Apply protection with the original password (oldPassword not required after unprotect)
            sheet.Protect(ProtectionType.All, password, string.Empty);

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

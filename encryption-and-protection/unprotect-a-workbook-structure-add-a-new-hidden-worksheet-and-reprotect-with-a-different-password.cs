// Title: Unprotect a workbook’s structure, add a hidden sheet, and re‑protect it with a new password using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to remove existing workbook structure protection with a known password, add a concealed worksheet, and then protect the structure again using a different password. | Write C# code that loads or creates an Excel workbook, calls Workbook.Unprotect, creates a concealed worksheet, and applies Workbook.Protect with ProtectionType.Structure and a new password.
// Common Searches: unprotect workbook structure and then hide a new sheet using Aspose.Cells in C# | protect only the workbook structure with a new password after adding a hidden sheet Aspose.Cells | hide a worksheet and set structure protection using Aspose.Cells .NET | change workbook protection password while keeping hidden sheets Aspose.Cells | unprotect workbook, add hidden sheet, re‑protect structure example Aspose.Cells
// Tags: unprotect workbook structure Aspose.Cells | insert hidden sheet Aspose.Cells C# | protect workbook structure with new password Aspose.Cells | excel structure protection Aspose.Cells | hide worksheet programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // C# example that unprotects an Excel workbook's structure with the current password, adds a hidden worksheet named "HiddenSheet", re‑applies structure protection using a new password, and saves the file as output.xlsx using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook. To load an existing file, uncomment the lines below
                // string inputPath = "input.xlsx";
                // Workbook workbook;
                // if (File.Exists(inputPath))
                // {
                //     workbook = new Workbook(inputPath);
                // }
                // else
                // {
                //     Console.WriteLine($"Input file not found: {inputPath}");
                //     return;
                // }
                Workbook workbook = new Workbook(); // create a new workbook

                // Unprotect the workbook structure if it is protected
                // If the workbook is not protected, this call has no effect
                workbook.Unprotect("CurrentPassword");

                // Add a new worksheet and hide it
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet newSheet = workbook.Worksheets[newSheetIndex];
                newSheet.Name = "HiddenSheet";
                newSheet.IsVisible = false; // hide the sheet

                // Protect only the workbook structure with a new password
                workbook.Protect(ProtectionType.Structure, "NewPassword");

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

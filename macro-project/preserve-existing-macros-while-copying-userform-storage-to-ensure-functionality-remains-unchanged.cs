// Title: Copy an XLSM workbook with VBA UserForms while preserving macros using Aspose.Cells for .NET
// Common Searches: how to duplicate a macro-enabled .xlsm file with UserForms using Aspose.Cells in C# | preserve VBA UserForm storage when saving a workbook as Xlsm with Aspose.Cells .NET | copy Excel workbook with macros and keep macro functionality using Aspose.Cells C# | save workbook to new location without losing macros Aspose.Cells SaveFormat.Xlsm | Aspose.Cells copy workbook while retaining VBA code and UserForms
// Tags: copy Xlsm workbook Aspose.Cells | retain VBA UserForms .NET | save as Xlsm retaining macros | workbook duplication with macro support | Aspose.Cells macro-enabled file handling

using System;
using System.IO;
using Aspose.Cells;

// The example loads a macro‑enabled .xlsm workbook that contains VBA UserForms, ensures the output directory exists, and saves it to a new location using SaveFormat.Xlsm. Aspose.Cells automatically retains all VBA code and UserForm storage, providing a straightforward way to duplicate macro‑enabled workbooks without losing functionality.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook containing macros and UserForms
            string sourcePath = @"C:\Input\SourceWithMacros.xlsm";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook (macros are retained automatically)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Path for the destination workbook
            string destinationPath = @"C:\Output\DestinationWithMacros.xlsm";

            // Ensure the output directory exists
            string? destDir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save the workbook to the destination path preserving macros
            sourceWorkbook.Save(destinationPath, SaveFormat.Xlsm);
            Console.WriteLine($"Destination workbook saved successfully: {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

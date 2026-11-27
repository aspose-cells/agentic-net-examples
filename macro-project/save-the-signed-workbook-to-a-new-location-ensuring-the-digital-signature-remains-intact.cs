// Title: Copy a digitally signed Excel workbook to a new folder while preserving its signature using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save a signed Excel file to another path without losing the signature | how to copy an .xlsx that has a digital signature using Aspose.Cells API | duplicate a digitally signed workbook while keeping the signature intact in .NET | C# example for preserving digital signature when moving an Excel workbook with Aspose.Cells | save signed workbook to new location preserving signature Aspose.Cells
// Tags: Aspose.Cells copy signed workbook | SaveFormat.Xlsx digital signature handling | Aspose.Cells signed workbook export | retain workbook signature on save | C# preserve signed Excel file with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a digitally signed Excel workbook, ensures the target directory exists, and saves the workbook to a new location using SaveFormat.Xlsx. The operation keeps the original digital signature intact and includes basic file existence checks and exception handling.
class Program
{
    static void Main()
    {
        // Path to the signed workbook
        string sourcePath = @"C:\Input\SignedWorkbook.xlsx";

        // Destination path where the signed workbook will be saved
        string destinationPath = @"C:\Output\SignedWorkbook_Copy.xlsx";

        try
        {
            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the signed workbook
            Workbook workbook = new Workbook(sourcePath);

            // Ensure the destination directory exists
            string destDir = Path.GetDirectoryName(destinationPath);
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Save the workbook to the new location, preserving the digital signature
            workbook.Save(destinationPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook copied successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: How to unprotect an Excel workbook’s structure, rename a worksheet, and re‑protect the structure with a new password using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, calls Unprotect with the current password, changes the name of the first worksheet, then calls Protect with ProtectionType.Structure and a new password, and saves the workbook. | Show how to programmatically remove workbook structure protection, rename a sheet, and apply a new structure password in Aspose.Cells for .NET, including error handling for missing input files.
// Common Searches: Aspose.Cells C# unprotect workbook structure then rename sheet | change worksheet name after removing structure protection using Aspose.Cells | apply new structure password to Excel file with Aspose.Cells .NET example
// Tags: unprotect workbook structure Aspose.Cells C# | rename worksheet after unprotect Aspose.Cells | protect workbook structure with new password Aspose.Cells | Aspose.Cells workbook structure protection example | C# change sheet name in protected workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads input.xlsx, removes existing workbook structure protection using the old password, renames the first worksheet to "RenamedSheet", re‑applies structure protection with a new password, and saves the result as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.xlsx";
                const string oldPassword = "oldPassword";
                const string newPassword = "newPassword";

                // Ensure the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Unprotect the workbook structure using the old password (if any)
                workbook.Unprotect(oldPassword);

                // Rename the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "RenamedSheet";

                // Protect only the workbook structure with a new password
                workbook.Protect(ProtectionType.Structure, newPassword);

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

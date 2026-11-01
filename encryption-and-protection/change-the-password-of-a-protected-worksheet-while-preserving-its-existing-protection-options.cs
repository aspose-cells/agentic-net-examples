// Title: Change a protected worksheet’s password while retaining all protection options with Aspose.Cells for .NET (C#)
// Common Searches: how to update worksheet password in Aspose.Cells C# without resetting protection settings | change Excel sheet protection password programmatically using Aspose.Cells .NET | preserve worksheet protection options when changing password with Aspose.Cells | unprotect and re‑protect a worksheet with a new password using Aspose.Cells API | replace old worksheet password with new one while keeping protection flags in C#
// Tags: worksheet password change Aspose.Cells | retain worksheet protection settings C# | unprotect worksheet old password Aspose.Cells | protect worksheet with new password all options .NET | Excel worksheet protection password update Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, unprotects the first worksheet using the current password, then re‑applies protection with the same options but a new password, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string oldPassword = "oldPassword";
            const string newPassword = "newPassword";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Unprotect the worksheet using the old password
                sheet.Unprotect(oldPassword);

                // Protect the worksheet again with the new password (all options)
                sheet.Protect(ProtectionType.All, newPassword, null);

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

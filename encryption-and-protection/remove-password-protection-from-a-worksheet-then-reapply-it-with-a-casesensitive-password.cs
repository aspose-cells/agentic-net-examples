// Title: How to remove worksheet protection and apply a case‑sensitive password with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unprotect worksheet and protect with case‑sensitive password | remove existing protection from Excel sheet then apply new password using Aspose.Cells | Aspose.Cells protect worksheet with ProtectionType.All and custom password in .NET | change worksheet password programmatically with Aspose.Cells in C# | unprotect Excel worksheet before re‑applying protection in C#
// Tags: Aspose.Cells worksheet unprotect API | Aspose.Cells protect worksheet with password | C# Excel worksheet case‑sensitive protection | Aspose.Cells ProtectionType.All usage | modify worksheet protection programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, clears any existing protection from the first worksheet, then reapplies protection using ProtectionType.All with a case‑sensitive password, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string password = "MySecretPassword";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify by name/index)
                Worksheet sheet = workbook.Worksheets[0];

                // Remove any existing protection from the worksheet
                sheet.Unprotect(string.Empty);

                // Re‑apply protection with a case‑sensitive password
                // The overload requires protection type, new password, and old password (empty if none)
                sheet.Protect(ProtectionType.All, password, string.Empty);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

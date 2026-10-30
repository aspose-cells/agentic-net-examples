// Title: Protect an unprotected Excel workbook with a strong password using Aspose.Cells Protect method in C#
// Common Searches: C# Aspose.Cells protect workbook with password only if not already protected | How to add full workbook protection using Aspose.Cells Protect in .NET | Set ProtectionType.All and password when saving an Excel file with Aspose.Cells | Check file existence and apply password protection to Excel using Aspose.Cells C# | Apply strong password to an Excel workbook programmatically with Aspose.Cells
// Tags: Aspose.Cells workbook protection API | ProtectionType.All usage Aspose.Cells | Excel workbook password encryption Aspose.Cells | File existence check before loading Aspose.Cells | Save workbook with protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies the input Excel file, loads it into an Aspose.Cells Workbook, applies full workbook protection with a strong password using the Protect method (ProtectionType.All), saves the protected workbook to a new file, and handles any runtime exceptions gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string password = "Str0ngP@ssw0rd!#2026";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Protect the entire workbook with the specified password
                // (Aspose.Cells will overwrite existing protection if present)
                workbook.Protect(ProtectionType.All, password);

                // Save the protected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

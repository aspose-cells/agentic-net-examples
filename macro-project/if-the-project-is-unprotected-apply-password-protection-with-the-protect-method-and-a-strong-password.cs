// Title: Apply full workbook password protection with a strong password using Aspose.Cells Protect method in C#
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, applies Workbook.Protect with ProtectionType.All and a strong password, then saves to a new file. | Modify the example to first detect if the workbook is already protected and only set a new password when it is unprotected. | Create a version that protects only the workbook structure and windows (not cells) using Protect with specific ProtectionType flags and a custom password.
// Common Searches: how to use Aspose.Cells Protect method to secure an entire Excel workbook in C# | C# example for applying a strong password to an unprotected .xlsx file with Aspose.Cells | Aspose.Cells set ProtectionType.All password protection code snippet | save a password‑protected workbook to a different location using Aspose.Cells | check workbook protection status before applying password with Aspose.Cells .NET
// Tags: Aspose.Cells Workbook.Protect method C# | full workbook password protection Excel Aspose.Cells | ProtectionType.All usage Aspose.Cells | save protected workbook to new file Aspose.Cells | detect existing workbook protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an existing Excel workbook, applies full protection with a strong password via Workbook.Protect(ProtectionType.All), and saves the protected file to a specified output path.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Apply password protection (all protection types)
                // If the workbook is already protected, this will overwrite the existing protection.
                workbook.Protect(ProtectionType.All, "Str0ngP@ssw0rd!2026");

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the protected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

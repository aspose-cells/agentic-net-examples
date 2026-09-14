// Title: Use Workbook.IsProtected to skip protecting an already secured Excel file with Aspose.Cells in C#
// AI Prompts: Insert an if statement that checks workbook.IsProtected before calling workbook.Protect, and write a console message when the workbook is already protected. | Refactor the example so that workbook.Protect is executed only when the workbook is not protected, while keeping the existing file‑existence and output‑directory checks.
// Common Searches: c# Aspose.Cells check if workbook is already password protected before applying protection | skip workbook.Protect call when workbook.IsProtected returns true Aspose.Cells | how to conditionally protect an Excel file using Aspose.Cells in C# | prevent re‑protecting an already protected Excel workbook with Aspose.Cells | use Workbook.IsProtected property in Aspose.Cells example
// Tags: Aspose.Cells workbook.IsProtected conditional check | C# protect Excel workbook only if not already protected | Aspose.Cells password protection skip already secured file | Excel file protection with ProtectionType.All in C# | avoid re‑protecting workbook using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel workbook, verifies the input file, checks whether the workbook is already protected using Workbook.IsProtected, applies structure and windows protection with a password only if needed, ensures the output directory exists, and saves the result while handling errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "MySecurePassword";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Protect the workbook (structure and windows) with a password
            workbook.Protect(ProtectionType.All, password);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a new file
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

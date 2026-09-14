// Title: Create a password‑protected Excel workbook that allows view‑only access and blocks edits using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an existing .xlsx file, applies full workbook protection with a password, and saves the file so it can be opened read‑only without allowing any edits. | Show how to use Aspose.Cells' Workbook.Protect method to enforce view‑only mode on an Excel workbook and prevent users from saving changes after opening it.
// Common Searches: Aspose.Cells C# protect entire workbook with password to make it read‑only | How to disable editing and saving in an Excel file using Aspose.Cells .NET | Create view‑only Excel workbook programmatically with Aspose.Cells and password protection
// Tags: Aspose.Cells Workbook.Protect password protection | C# read‑only Excel workbook generation | prevent Excel edits with Aspose.Cells | save protected workbook as .xlsx using Aspose | view‑only Excel file Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookProtectionExample
{
    // The example loads input.xlsx, applies full workbook protection with the password "protect123" via Workbook.Protect(ProtectionType.All, password), and saves the protected file as protected_output.xlsx, ensuring the workbook can be opened for viewing only.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "protected_output.xlsx";
            const string password = "protect123";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Protect the entire workbook with a password
                workbook.Protect(ProtectionType.All, password);

                // Save the protected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Protected workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

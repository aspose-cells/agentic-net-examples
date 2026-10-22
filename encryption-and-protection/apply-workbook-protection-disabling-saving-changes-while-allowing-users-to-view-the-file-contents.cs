// Title: Apply password-protected write protection to an Excel workbook to prevent saving changes while still allowing view using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# prevent users from saving changes to an Excel workbook | Make an Excel file read-only with password using Aspose.Cells .NET | How to apply write protection to a workbook without blocking view in C# | Example code for Aspose.Cells Protect method to disable editing | Set workbook protection type All in Aspose.Cells to create view-only Excel file
// Tags: Aspose.Cells workbook write protection C# | Excel read-only protection with password Aspose .NET | Prevent saving changes in Excel using Aspose.Cells Protect | Aspose.Cells Protect method example C# | Password protected view-only workbook Aspose.Cells | Aspose.Cells ProtectionType.All usage

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input file, loads it into a Workbook, applies password‑protected write protection using workbook.Protect(ProtectionType.All, password) to make the file view‑only, saves the protected workbook, and handles any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "protected_output.xlsx";
        const string password = "MyPassword";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply write protection with a password
            workbook.Protect(ProtectionType.All, password);

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

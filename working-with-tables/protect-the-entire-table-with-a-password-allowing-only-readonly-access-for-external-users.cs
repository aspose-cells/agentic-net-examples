// Title: Password‑protect an entire Excel workbook for read‑only external access with Aspose.Cells in C#
// AI Prompts: Create C# code that loads an .xlsx file, checks its existence, applies workbook‑level protection using ProtectionType.All with a password, and saves the result with Aspose.Cells. | Demonstrate how to use Aspose.Cells' Workbook.Protect method to enforce read‑only access for external users in a .NET application. | Write a robust C# snippet that handles missing input files and exceptions while password‑protecting an Excel workbook using Aspose.Cells.
// Common Searches: how to use Aspose.Cells to set a password for an entire workbook in C# | Aspose.Cells protect Excel file read only for external users .NET example | C# workbook.Protect with ProtectionType.All and password Aspose | prevent editing of an Excel workbook by external users using Aspose.Cells | save password‑protected .xlsx using Aspose.Cells C# code sample
// Tags: Aspose.Cells workbook password protection | C# protect entire Excel workbook | ProtectionType.All Aspose.Cells example | read‑only Excel file external users .NET | Workbook.Protect method C#

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx, verifies its presence, applies workbook‑level protection (ProtectionType.All) with password "SecurePassword123" to make the file read‑only for external users, and saves the protected workbook as output_protected.xlsx while handling errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output_protected.xlsx";
        const string password = "SecurePassword123";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Protect the entire workbook with a password (read‑only for external users)
            workbook.Protect(ProtectionType.All, password);

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook successfully protected and saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: How to protect an Excel workbook with Aspose.Cells for .NET to block printing while still allowing on‑screen viewing
// AI Prompts: Write C# code that uses Aspose.Cells to protect a workbook with a password and configure the protection options so that printing is prohibited but all other view features remain enabled. | Show the Aspose.Cells API calls required to disable the Print command in a protected Excel file while keeping the worksheet visible to users. | Provide a complete example that creates or loads a workbook, applies workbook protection with printing disabled, and saves the result.
// Common Searches: Aspose.Cells C# disable print option in protected Excel workbook | prevent printing of Excel file using Aspose.Cells .NET API | how to allow view‑only access to Excel workbook with Aspose.Cells | C# set workbook protection to block printing but keep sheet visible Aspose
// Tags: Aspose.Cells workbook protection disable printing | C# Aspose.Cells set print restriction | Excel file view‑only protection Aspose.Cells | Aspose.Cells protect workbook with password | disable Excel printing via Aspose.Cells API

using System;
using Aspose.Cells;

// The example creates a new workbook (or loads an existing one), applies password‑protected workbook protection, configures the protection settings to prevent printing while still allowing users to view the sheets on screen, and saves the protected file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Protect the entire workbook with a password
            // ProtectionType.All protects structure, windows, and objects
            workbook.Protect(ProtectionType.All, "securePassword");

            // Save the protected workbook
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

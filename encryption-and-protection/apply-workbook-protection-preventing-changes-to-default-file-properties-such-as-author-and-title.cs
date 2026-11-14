// Title: How to add password‑protected structure protection to an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to protect Excel workbook structure with a password using Aspose.Cells | Aspose.Cells Workbook.Protect example for preventing sheet addition or deletion | Saving a password‑protected .xlsx file with Aspose.Cells in a .NET application | Apply workbook structure protection without affecting document properties in Aspose.Cells | Protecting Excel file against worksheet renaming using Aspose.Cells .NET API
// Tags: Aspose.Cells workbook structure protection C# | password protect Excel workbook Aspose.Cells .NET | Workbook.Protect method example | prevent worksheet addition deletion Aspose.Cells | save protected .xlsx with Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates creating or loading a workbook, applying password‑protected structure protection via Workbook.Protect(ProtectionType.Structure, password), and saving the workbook as a protected .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // NOTE: In some versions of Aspose.Cells the DefaultFileProperties
            // property is not available. If needed, document properties can be
            // set via the built‑in file properties after saving the file.
            // The following lines are omitted to keep the code compatible
            // with all supported versions.

            // Apply workbook protection with a password.
            // The Structure protection prevents changes to the workbook
            // structure (e.g., adding, removing, or renaming worksheets).
            workbook.Protect(ProtectionType.Structure, "StrongPassword123");

            // Save the protected workbook
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

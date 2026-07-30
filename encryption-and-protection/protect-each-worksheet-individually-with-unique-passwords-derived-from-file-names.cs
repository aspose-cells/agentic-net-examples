// Title: Aspose.Cells .NET – Protect Each Worksheet with a Unique Password Derived from the File Name
// Description: Loads an Excel workbook, uses the file name (without extension) as a base, generates a distinct password for every sheet by appending the sheet index, applies full protection, and saves the protected file. Demonstrates C# worksheet protection with Aspose.Cells.
// Keywords: Aspose.Cells worksheet protection | C# protect Excel sheet with password | unique password per worksheet | derive password from file name | Aspose.Cells .NET example | Excel security programmatically | GitHub Aspose.Cells sample | global Excel encryption
// Common Searches: protect each Excel worksheet with a different password using Aspose.Cells | generate worksheet passwords from workbook file name C# | Aspose.Cells protect all sheets programmatically | C# example for per‑sheet password protection in Excel | how to lock individual sheets with unique passwords Aspose.Cells
// Developer Intent: Apply a separate password to every worksheet, using a password that incorporates the workbook’s file name for traceability.
// Use Cases: Distribute a multi‑sheet report where each sheet is accessible only with its own password. | Automate compliance documents that require sheet‑level security tied to the source file name. | Create tenant‑specific worksheets in a single workbook, each locked with a unique password derived from the report name.
// AI Prompts: Generate C# code with Aspose.Cells that protects each worksheet using a password composed of the workbook name and sheet index. | Show how to add a custom salt or timestamp to the generated worksheet passwords instead of the sheet index. | Explain the steps to unprotect a specific worksheet when the generated password is known, using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, uses the file name (without extension) as a base, generates a distinct password for every sheet by appending the sheet index, applies full protection, and saves the protected file. Demonstrates C# worksheet protection with Aspose.Cells.
class ProtectWorksheets
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path where the protected workbook will be saved
        string outputPath = "output_protected.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Derive a base password from the file name (without extension)
        string basePassword = Path.GetFileNameWithoutExtension(inputPath);

        // Iterate through each worksheet and apply a unique password
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Create a unique password for the current worksheet
            string sheetPassword = $"{basePassword}_Sheet{sheet.Index}";

            // Protect the worksheet with all protection types and the generated password
            // Using the Protect(ProtectionType, string, string) overload as per the rule set
            sheet.Protect(ProtectionType.All, sheetPassword, null);
        }

        // Save the workbook (lifecycle rule: save)
        workbook.Save(outputPath);
    }
}

// Title: Check worksheet protection status and type with Aspose.Cells for .NET
// Description: Shows how to detect if a worksheet is protected, whether the protection is password‑based (read‑only), and how to read the applied ProtectionType using Aspose.Cells in C#.
// Keywords: Aspose.Cells worksheet protection | Worksheet.IsProtected | Protection.IsProtectedWithPassword | ProtectionType enum | .NET Excel security | C# Aspose.Cells example
// Common Searches: aspocells check if worksheet is protected | detect password protection on Excel sheet using Aspose.Cells | get protection type of a worksheet in C# | read‑only worksheet protection Aspose.Cells | how to verify worksheet protection after loading workbook
// Developer Intent: Determine whether a worksheet is protected, confirm if the protection uses a password (read‑only), and retrieve the specific ProtectionType applied.
// Use Cases: Validate protection before allowing edits in a web or desktop application. | Audit Excel files for compliance by confirming password‑based read‑only protection. | Log or display the protection mode (e.g., ProtectionType.All) for user awareness.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells and returns a tuple (bool isProtected, bool isPasswordProtected, ProtectionType type) for the first worksheet. | Write a method using Aspose.Cells that writes the worksheet protection status and ProtectionType to a log file instead of the console. | Create a reusable utility class in C# that checks any worksheet's protection state and exposes properties for IsProtected, IsPasswordProtected, and ProtectionType.

using System;
using Aspose.Cells;

// Shows how to detect if a worksheet is protected, whether the protection is password‑based (read‑only), and how to read the applied ProtectionType using Aspose.Cells in C#.
class WorksheetProtectionCheck
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet with a password (read‑only style protection)
        sheet.Protect(ProtectionType.All, "pwd123", null);

        // Save the workbook
        string filePath = "protectedWorksheet.xlsx";
        workbook.Save(filePath);

        // Load the saved workbook
        Workbook loadedWorkbook = new Workbook(filePath);
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // Check if the worksheet is protected
        bool isProtected = loadedSheet.IsProtected;

        // Check if the protection is password‑based (read‑only)
        bool isPasswordProtected = loadedSheet.Protection.IsProtectedWithPassword;

        // Report the protection status and type
        Console.WriteLine($"Worksheet is protected: {isProtected}");
        Console.WriteLine($"Protected with password (read‑only): {isPasswordProtected}");
        Console.WriteLine($"Protection type used: {ProtectionType.All}");
    }
}

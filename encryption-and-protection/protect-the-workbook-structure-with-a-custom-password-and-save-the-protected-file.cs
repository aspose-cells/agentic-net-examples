// Title: How to protect an Excel workbook's structure with a custom password and save it as XLSX using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to apply structure protection with a specified password to a workbook and then save it as an XLSX file. | Show how to lock sheet addition, deletion, hiding, and reordering by calling Workbook.Protect with ProtectionType.Structure and a custom password, then persist the workbook.
// Common Searches: Aspose.Cells C# protect workbook structure with password example | C# code to prevent adding or deleting sheets in Excel using Aspose.Cells | Save a password‑protected XLSX file with Aspose.Cells .NET | Workbook.Protect ProtectionType.Structure usage in Aspose.Cells
// Tags: Aspose.Cells workbook.Protect structure password | C# protect Excel workbook structure | save password‑protected XLSX Aspose.Cells | ProtectionType.Structure usage Aspose.Cells | custom password Excel protection .NET

using Aspose.Cells;

// The example creates a new Workbook (or loads an existing one), applies structure protection with the password "MyCustomPassword" using ProtectionType.Structure, and saves the protected workbook as "ProtectedWorkbook.xlsx" in XLSX format.
class Program
{
    static void Main()
    {
        // Create a new workbook (you can also load an existing one with new Workbook("input.xlsx"))
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a custom password
        // ProtectionType.Structure ensures that sheets cannot be added, deleted, hidden, or reordered
        workbook.Protect(ProtectionType.Structure, "MyCustomPassword");

        // Save the protected workbook to a file
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}

// Title: Protect Workbook Structure with Password in Aspose.Cells for .NET (C#) – Block New Sheets, Allow Renaming
// Description: C# example that creates a workbook, applies structure protection using Workbook.Protect(ProtectionType.Structure, password), prevents insertion of additional worksheets while still permitting programmatic renaming of existing sheets, and saves the file as ProtectedWorkbook.xlsx.
// Keywords: Aspose.Cells workbook protection | C# protect workbook structure | prevent adding worksheets Excel | allow sheet rename after protection | Excel password protection .NET | Workbook.Protect Structure | Aspose.Cells encryption and protection
// Common Searches: Aspose.Cells protect workbook structure C# | block new worksheets but rename existing sheets | Excel workbook password protection .NET example | how to prevent adding sheets in Aspose.Cells | rename sheets after workbook protection
// Developer Intent: Apply structure-level protection to an Excel workbook so users cannot add new worksheets but can rename the ones that already exist.
// Use Cases: Distribute a template where the sheet layout must stay fixed while users customize sheet names. | Secure a financial model that forbids extra worksheets but allows department‑specific renaming. | Create a localized report workbook that locks sheet count and order yet permits language‑specific sheet titles.
// AI Prompts: Generate C# code that protects only the workbook structure with a password using Aspose.Cells and demonstrates that adding a new worksheet fails. | Show how to unprotect a structure‑protected workbook with the same password and then rename a sheet in Aspose.Cells for .NET. | Explain how to combine workbook structure protection with individual worksheet cell protection in Aspose.Cells.

using Aspose.Cells;

// C# example that creates a workbook, applies structure protection using Workbook.Protect(ProtectionType.Structure, password), prevents insertion of additional worksheets while still permitting programmatic renaming of existing sheets, and saves the file as ProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password.
        // This blocks the creation of new worksheets while still permitting
        // existing worksheets to be renamed programmatically.
        workbook.Protect(ProtectionType.Structure, "myPassword");

        // Save the protected workbook
        workbook.Save("ProtectedWorkbook.xlsx");
    }
}

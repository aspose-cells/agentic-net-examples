// Title: Apply password‑protected write protection to an Excel workbook and then remove it programmatically with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect entire workbook with a password and save | How to unprotect a password‑protected Excel file using Aspose.Cells in .NET | Remove write protection from an Excel workbook using the same password with Aspose.Cells | Aspose.Cells Protect method with ProtectionType.All example in C# | Programmatically toggle workbook protection in Aspose.Cells .NET
// Tags: Aspose.Cells Protect method password | Aspose.Cells Unprotect workbook C# | write protection Excel workbook Aspose.Cells | protect entire workbook structure Aspose.Cells | remove Excel write protection programmatically .NET

using Aspose.Cells;

// The example creates a new Workbook, applies write protection for the whole workbook using Protect(ProtectionType.All, "SecretPassword"), saves it, reloads the protected file, removes the protection with Unprotect("SecretPassword"), and saves the unprotected workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Apply write protection with a password
        // ProtectionType.All protects the entire workbook (structure, windows, and write protection)
        workbook.Protect(ProtectionType.All, "SecretPassword");

        // Save the protected workbook
        workbook.Save("ProtectedWorkbook.xlsx");

        // Load the protected workbook
        Workbook protectedWb = new Workbook("ProtectedWorkbook.xlsx");

        // Remove protection using the same password
        protectedWb.Unprotect("SecretPassword");

        // Save the unprotected workbook
        protectedWb.Save("UnprotectedWorkbook.xlsx");
    }
}

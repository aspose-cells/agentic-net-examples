// Title: Encrypt an Excel workbook with a custom password and save it as an XLS (Excel 97‑2003) file using Aspose.Cells for .NET
// Common Searches: how to set a password on an Excel workbook using Aspose.Cells C# | C# encrypt an XLSX file and save as XLS with Aspose.Cells | Aspose.Cells .NET protect Excel file with user‑defined password | save password‑protected Excel 97‑2003 file in C# | apply workbook password before saving as XLS using Aspose.Cells
// Tags: Aspose.Cells workbook password encryption | C# save encrypted Excel 97-2003 file | Workbook.Settings.Password usage | convert XLSX to password‑protected XLS | Excel file protection with Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example loads an existing XLSX workbook, assigns a user‑specified password via the Workbook.Settings.Password property, and saves the workbook as an encrypted XLS (Excel 97‑2003) file using SaveFormat.Excel97To2003.
class Program
{
    static void Main()
    {
        // Path to the source workbook (plain Excel file)
        string inputPath = "input.xlsx";

        // Path where the encrypted workbook will be saved (XLS format)
        string outputPath = "encrypted.xls";

        // User‑specified password for encryption
        string password = "YourPasswordHere";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Apply password protection (encryption)
        workbook.Settings.Password = password;

        // Save the workbook as an XLS file (Excel 97‑2003 format)
        workbook.Save(outputPath, SaveFormat.Excel97To2003);
    }
}

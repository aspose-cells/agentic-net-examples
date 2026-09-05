// Title: Encrypt an Excel workbook with a custom password and save it as an XLS file using Aspose.Cells for .NET
// AI Prompts: Load a workbook, set Workbook.Settings.Password to a user‑provided string, and save it as an Excel97To2003 (.xls) file with Aspose.Cells. | Modify the example to read the password from console input, apply it to the workbook, and generate an encrypted XLS output. | Demonstrate how to programmatically protect an existing .xlsx file with a password and export the protected version as a legacy .xls using Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to password protect an Excel file and export to .xls format | C# encrypt workbook with user defined password using Aspose.Cells SaveFormat.Excel97To2003 | Set password on Workbook.Settings and save as encrypted XLS with Aspose.Cells example
// Tags: Workbook.Settings.Password encryption | encrypted XLS export using Aspose.Cells | Excel97To2003 password protection C# | protect workbook programmatically Aspose.Cells | export encrypted Excel 97-2003 file

using System;
using Aspose.Cells;

// // Loads a workbook, assigns a password via Workbook.Settings.Password, and saves it as an encrypted Excel 97‑2003 (.xls) file.
class Program
{
    static void Main()
    {
        // Path to the source workbook (can be any supported format)
        string sourcePath = "source.xlsx";

        // Path for the encrypted XLS file to be created
        string encryptedPath = "encrypted.xls";

        // User‑specified password for encryption
        string password = "YourPasswordHere";

        // Load the existing workbook
        Workbook workbook = new Workbook(sourcePath);

        // Apply password protection (encryption) to the workbook
        workbook.Settings.Password = password;

        // Save the workbook as an Excel 97‑2003 file (XLS) with encryption
        workbook.Save(encryptedPath, SaveFormat.Excel97To2003);
    }
}

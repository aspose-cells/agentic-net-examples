// Title: Detect legacy XLS (Excel 97-2003) format and apply password protection with Aspose.Cells for .NET
// Common Searches: how to check if an Excel workbook is XLS format using Aspose.Cells C# | apply password protection to a binary .xls file with Aspose.Cells | save encrypted workbook in original format (XLS or XLSX) using Aspose.Cells | set Workbook.Settings.Password for legacy Excel files in C# | detect Excel97To2003 format before encrypting with Aspose.Cells
// Tags: identify legacy binary XLS using Aspose.Cells | apply password encryption to Excel97To2003 workbook | choose SaveFormat based on source file type Aspose.Cells | Workbook.Settings.Password example C# | encrypt Excel file while preserving original format

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, determines if it is the legacy Excel97To2003 (XLS) format via Workbook.FileFormat, sets a password with Workbook.Settings.Password, selects the appropriate SaveFormat (Excel97To2003 for .xls or Xlsx for .xlsx), and saves the encrypted file, preserving the original file type.
class ExcelEncryption
{
    static void Main()
    {
        // Paths for source and encrypted files
        string sourcePath = "input.xls";      // can be .xls or .xlsx
        string encryptedPath = "encrypted.xls";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook (Aspose.Cells automatically detects the format)
            Workbook workbook = new Workbook(sourcePath);

            // Determine if the workbook is in legacy binary format (XLS)
            bool isLegacyXls = workbook.FileFormat == FileFormatType.Excel97To2003;

            // Set password for encryption
            workbook.Settings.Password = "MySecretPassword";

            // Choose appropriate save format based on the original file type
            SaveFormat saveFormat = isLegacyXls ? SaveFormat.Excel97To2003 : SaveFormat.Xlsx;

            // Save the encrypted workbook
            workbook.Save(encryptedPath, saveFormat);

            Console.WriteLine($"File encrypted successfully. Legacy XLS: {isLegacyXls}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

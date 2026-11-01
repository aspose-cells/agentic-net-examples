// Title: Detect legacy .xls workbook, convert to .xlsx, and apply password protection using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells detect if Excel file is .xls and convert to .xlsx before encrypting | How to convert legacy XLS to XLSX in memory and set opening password with Aspose.Cells | Apply password protection to an Excel workbook after format conversion using Aspose.Cells .NET | Check file extension and save encrypted XLSX with Aspose.Cells in C# | Aspose.Cells encrypt workbook only after converting from .xls to .xlsx
// Tags: detect legacy XLS Aspose.Cells | convert XLS to XLSX using memory stream | apply opening password Aspose.Cells | encrypt workbook as XLSX SaveFormat | auto-detect Excel format Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, determines whether the source file is a legacy .xls format, converts it to .xlsx via a memory stream when needed, sets an opening password, and saves the encrypted workbook as an .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source workbook (could be .xls or .xlsx)
        string sourcePath = "input.xls";

        // Path for the final encrypted workbook (XLSX)
        string encryptedPath = "encrypted.xlsx";

        // Load the workbook; Aspose.Cells auto-detects the format
        Workbook workbook = new Workbook(sourcePath);

        // Determine if the original file is a legacy XLS (BIFF) format
        bool isLegacyXls = Path.GetExtension(sourcePath)
                               .Equals(".xls", StringComparison.OrdinalIgnoreCase);

        if (isLegacyXls)
        {
            // Convert legacy XLS to XLSX by saving to a memory stream in Xlsx format
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for reading
                workbook = new Workbook(ms); // Load the XLSX version from the stream
            }
        }

        // Apply encryption by setting a password for opening the workbook
        workbook.Settings.Password = "MySecretPassword";

        // Save the encrypted workbook as XLSX
        workbook.Save(encryptedPath, SaveFormat.Xlsx);
    }
}

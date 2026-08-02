// Title: Detect Deprecated Excel 97‑2003 Encryption with Aspose.Cells for .NET
// Description: Uses Aspose.Cells FileFormatUtil to identify encrypted XLS files that rely on the legacy Excel 97‑2003 encryption algorithm and flags them for migration to a modern format without fully loading the workbook.
// Keywords: Aspose.Cells encryption detection | detect encrypted XLS | legacy Excel 97-2003 encryption | deprecated encryption algorithm | FileFormatUtil C# | Excel workbook migration | modern encryption Aspose.Cells | C# Excel security check
// Common Searches: how to check if an XLS file uses old encryption with Aspose.Cells | detect legacy Excel encryption without opening workbook .NET | flag encrypted Excel 97‑2003 files for migration | Aspose.Cells detect encrypted workbook example | C# identify deprecated Excel encryption algorithm
// Developer Intent: Determine whether an Excel workbook is encrypted with a deprecated algorithm and mark it for upgrade.
// Use Cases: Batch‑scan a directory of .xls files, log those using legacy encryption, and schedule conversion to .xlsx. | Add a pre‑deployment gate in CI/CD pipelines that blocks workbooks encrypted with outdated algorithms. | Automatically load flagged workbooks, re‑save them with modern encryption, and archive the original files.
// AI Prompts: Generate a C# method that returns true if a given file path points to an XLS workbook encrypted with the deprecated Excel 97‑2003 algorithm using Aspose.Cells. | Create a script that iterates over all Excel files in a folder, detects deprecated encryption, and converts each to .xlsx with a strong password. | Write a PowerShell wrapper that calls a .NET utility to report encrypted legacy workbooks on a network share.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // Uses Aspose.Cells FileFormatUtil to identify encrypted XLS files that rely on the legacy Excel 97‑2003 encryption algorithm and flags them for migration to a modern format without fully loading the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be examined
            string workbookPath = "sample.xls"; // Change to your file path

            // Detect file format and encryption status without fully loading the workbook
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(workbookPath);

            // Determine if the workbook is encrypted
            bool isEncrypted = formatInfo.IsEncrypted;

            // Determine if the workbook uses the old Excel 97‑2003 format (XLS)
            bool isLegacyFormat = formatInfo.LoadFormat == LoadFormat.Excel97To2003;

            // Flag for migration if encrypted and using a legacy format (deprecated algorithms)
            if (isEncrypted && isLegacyFormat)
            {
                Console.WriteLine($"[ALERT] Workbook '{workbookPath}' is encrypted using a deprecated algorithm. Migration recommended.");
            }
            else if (isEncrypted)
            {
                Console.WriteLine($"Workbook '{workbookPath}' is encrypted with a modern algorithm.");
            }
            else
            {
                Console.WriteLine($"Workbook '{workbookPath}' is not encrypted.");
            }

            // Example of loading the workbook (if further processing is needed)
            // LoadOptions can include the password if known; here we assume no password for demonstration
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // (Optional) Save a copy after migration steps – placeholder for actual migration logic
            // workbook.Save("migrated_" + System.IO.Path.GetFileName(workbookPath), SaveFormat.Xlsx);
        }
    }
}

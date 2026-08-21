// Title: Detect Deprecated Excel Encryption (XOR/Compatible) with Aspose.Cells for .NET
// Description: Use Aspose.Cells' FileFormatUtil.DetectFileFormat to check if an Excel workbook is encrypted without loading it, then flag files that may rely on the obsolete XOR or Compatible encryption algorithms for migration to stronger protection.
// Keywords: Aspose.Cells encryption detection | deprecated Excel encryption | XOR encryption Excel | Compatible encryption Excel | FileFormatUtil DetectFileFormat | C# check workbook encryption | Excel security migration .NET | weak encryption upgrade Aspose
// Common Searches: How to detect deprecated XOR encryption in Excel with Aspose.Cells | Check if an XLSX file uses old Compatible encryption using C# | Aspose.Cells detect encrypted workbook without opening | Identify weak Excel encryption algorithms for migration | C# scan folder for Excel files with obsolete encryption
// Developer Intent: Determine whether an Excel workbook is encrypted and, if so, flag it for re‑encryption because it might be using the legacy XOR or Compatible algorithms.
// Use Cases: Quickly verify encryption status of incoming Excel files before processing them. | Generate alerts for security teams to re‑encrypt workbooks that rely on outdated algorithms. | Integrate into automated batch jobs that audit large collections of spreadsheets for compliance.
// AI Prompts: Write a C# routine that scans a directory, uses Aspose.Cells to detect encrypted Excel files, and lists those that could be using XOR or Compatible encryption. | Provide a method signature that returns true when a workbook is encrypted and suggests migration when the algorithm is deprecated. | Explain the steps to re‑encrypt an Excel file with AES‑256 using Aspose.Cells after detecting a deprecated encryption method.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // Use Aspose.Cells' FileFormatUtil.DetectFileFormat to check if an Excel workbook is encrypted without loading it, then flag files that may rely on the obsolete XOR or Compatible encryption algorithms for migration to stronger protection.
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be inspected
            string filePath = "workbook.xlsx";

            // Detect file format and encryption status without opening the workbook
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Flag if the workbook is encrypted
            if (formatInfo.IsEncrypted)
            {
                // NOTE: Aspose.Cells does not expose the exact encryption algorithm directly.
                // Deprecated algorithms are XOR and Compatible (Excel 97/2000).
                // If the workbook is encrypted, further analysis may be required to
                // determine the algorithm. Here we flag it for migration.
                Console.WriteLine($"[ALERT] The workbook \"{filePath}\" is encrypted. Verify if it uses a deprecated algorithm (XOR or Compatible) and migrate to a stronger encryption.");
            }
            else
            {
                Console.WriteLine($"The workbook \"{filePath}\" is not encrypted.");
            }
        }
    }
}

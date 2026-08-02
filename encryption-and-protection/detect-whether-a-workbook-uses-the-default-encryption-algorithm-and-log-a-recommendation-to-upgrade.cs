// Title: Detect Default Encryption in an Excel Workbook and Suggest AES Upgrade with Aspose.Cells for .NET
// Description: This C# example loads an .xlsx file using Aspose.Cells, checks the Workbook.Settings.IsDefaultEncrypted flag to see if the file is protected with the library's built‑in encryption, outputs a console notice, and advises switching to a stronger algorithm (such as AES) via SetEncryptionOptions before disposing the workbook.
// Keywords: Aspose.Cells default encryption | Workbook.Settings.IsDefaultEncrypted | Excel encryption strength .NET | upgrade to AES Aspose.Cells | check workbook protection | SetEncryptionOptions example | C# Excel security | detect weak encryption Excel
// Common Searches: Aspose.Cells how to know if workbook uses default encryption | C# check Excel file encryption level with Aspose | recommend stronger encryption for Excel using Aspose.Cells | property to detect built‑in encryption in Aspose.Cells | sample code to log encryption recommendation
// Developer Intent: Determine whether an Excel file is secured with the built‑in algorithm and provide guidance to apply a stronger cipher.
// Use Cases: Security audit of incoming spreadsheets in a regulated environment | Automated compliance logging for files that use weak encryption | Pre‑processing step in a data‑ingestion pipeline that re‑encrypts with AES‑256 | Batch utility to scan and flag workbooks before archival
// AI Prompts: Write C# code that reads an Excel workbook with Aspose.Cells, detects default encryption, and re‑encrypts it using AES‑256 via SetEncryptionOptions. | Explain how to integrate a built‑in encryption check into an existing Aspose.Cells file‑import workflow and log a compliance warning. | Create a script that iterates over a directory of .xlsx files, reports those using the library's default cipher, and upgrades them to AES‑256.

using System;
using Aspose.Cells;

namespace WorkbookEncryptionCheck
{
    // This C# example loads an .xlsx file using Aspose.Cells, checks the Workbook.Settings.IsDefaultEncrypted flag to see if the file is protected with the library's built‑in encryption, outputs a console notice, and advises switching to a stronger algorithm (such as AES) via SetEncryptionOptions before disposing the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be inspected
            string filePath = "sample.xlsx";

            // Load the workbook (no password needed for unencrypted files)
            Workbook workbook = new Workbook(filePath);

            // Check if the workbook is encrypted with the default algorithm
            bool isDefaultEncrypted = workbook.Settings.IsDefaultEncrypted;

            if (isDefaultEncrypted)
            {
                Console.WriteLine("The workbook uses the default encryption algorithm.");
                Console.WriteLine("Recommendation: Upgrade to a stronger encryption algorithm (e.g., AES) using SetEncryptionOptions.");
            }
            else
            {
                Console.WriteLine("The workbook does not use the default encryption algorithm.");
            }

            // Optional: Dispose the workbook to release resources
            workbook.Dispose();
        }
    }
}

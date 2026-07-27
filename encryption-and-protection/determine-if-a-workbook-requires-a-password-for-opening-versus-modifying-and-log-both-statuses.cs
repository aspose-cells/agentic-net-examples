// Title: Check if an Excel workbook needs an open or modify password with Aspose.Cells for .NET
// Description: Loads an Excel file using Aspose.Cells, determines whether the workbook is encrypted (open password) via Workbook.Settings.IsEncrypted and whether it is write‑protected (modify password) via Workbook.Settings.WriteProtection.IsWriteProtected, then prints both boolean results.
// Keywords: Aspose.Cells | C# password protection | Workbook.IsEncrypted | WriteProtection.IsWriteProtected | detect Excel encryption .NET | open password check | modify password check | Excel file security | Aspose.Cells encryption detection
// Common Searches: Aspose.Cells how to know if Excel file has an open password | C# check write protection on workbook with Aspose.Cells | determine if Excel workbook is encrypted using Aspose | get password status of Excel file Aspose.Cells | read workbook protection flags .NET
// Developer Intent: Find out whether a loaded workbook requires a password to open or to modify.
// Use Cases: Audit spreadsheets for protection before processing them. | Prompt users for a password only when the file is encrypted or write‑protected. | Exclude write‑protected workbooks from batch write operations. | Generate a security‑status report for a collection of Excel files.
// AI Prompts: Write C# code that uses Aspose.Cells to report open‑password and modify‑password requirements for a given Excel file, handling missing files and exceptions. | Explain the difference between Workbook.Settings.IsEncrypted and Workbook.Settings.WriteProtection.IsWriteProtected, and show typical usage scenarios in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file using Aspose.Cells, determines whether the workbook is encrypted (open password) via Workbook.Settings.IsEncrypted and whether it is write‑protected (modify password) via Workbook.Settings.WriteProtection.IsWriteProtected, then prints both boolean results.
    public class WorkbookPasswordStatusDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the workbook to be examined
            string filePath = "sample.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook within a using block to ensure proper disposal
            using (Workbook workbook = new Workbook(filePath))
            {
                // Check if a password is required to open the workbook
                bool requiresOpenPassword = workbook.Settings.IsEncrypted;

                // Check if a password is required to modify (write-protect) the workbook
                bool requiresModifyPassword = workbook.Settings.WriteProtection.IsWriteProtected;

                // Log the statuses
                Console.WriteLine($"Requires password to open: {requiresOpenPassword}");
                Console.WriteLine($"Requires password to modify: {requiresModifyPassword}");
            }
        }
    }
}

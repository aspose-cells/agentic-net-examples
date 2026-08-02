// Title: Save a Digitally Signed Excel Workbook as an Audit Copy Using Aspose.Cells for .NET
// Description: This sample loads an Excel file that includes a digital signature, validates the signature via the IsDigitallySigned property, and writes the workbook to a different filename. The source file stays untouched, creating an immutable audit trail while keeping the signature intact. Basic handling for missing files and runtime errors is also shown.
// Keywords: Aspose.Cells | C# | .NET | digital signature | Excel workbook | audit copy | preserve signature | Workbook.Save | file backup | IsDigitallySigned | error handling | file integrity
// Common Searches: Aspose.Cells save signed workbook C# | create audit copy of signed Excel file .NET | preserve digital signature when copying Excel | how to verify workbook signature Aspose.Cells | C# load and re‑save digitally signed Excel
// Developer Intent: Generate a read‑only backup of a signed Excel document without modifying the original.
// Use Cases: Archiving financial statements that must retain their original digital signature. | Automating compliance‑driven duplication of signed reports to a secure repository. | Implementing a version‑controlled audit log for contract documents stored as Excel files. | Providing a safe way to distribute signed spreadsheets to external partners while keeping the source file intact.
// AI Prompts: Write C# code that loads an Excel workbook, checks for a digital signature using Aspose.Cells, and saves it under a new name without losing the signature. | Explain how Aspose.Cells determines the output format when calling Workbook.Save with a filename. | Show error‑handling patterns for missing source files and signature verification failures in Aspose.Cells examples. | Describe best practices for creating immutable audit copies of signed Excel files in .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This sample loads an Excel file that includes a digital signature, validates the signature via the IsDigitallySigned property, and writes the workbook to a different filename. The source file stays untouched, creating an immutable audit trail while keeping the signature intact. Basic handling for missing files and runtime errors is also shown.
    public class SaveSignedWorkbookDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the original signed workbook
            string originalPath = "SignedOriginal.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Error: The file '{originalPath}' was not found.");
                return;
            }

            try
            {
                // Load the signed workbook
                Workbook signedWorkbook = new Workbook(originalPath);

                // Verify that the workbook is digitally signed
                bool isSigned = signedWorkbook.IsDigitallySigned;
                Console.WriteLine($"Original workbook is digitally signed: {isSigned}");

                // Define a new filename for the audit copy
                string auditCopyPath = "SignedAuditCopy.xlsx";

                // Save the workbook under the new filename.
                // The Save(string) overload automatically selects the format based on the file extension.
                signedWorkbook.Save(auditCopyPath);

                Console.WriteLine($"Signed workbook saved as audit copy: {auditCopyPath}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

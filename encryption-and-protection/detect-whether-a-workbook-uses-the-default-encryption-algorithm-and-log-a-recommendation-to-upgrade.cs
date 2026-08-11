// Title: Detect Aspose.Cells default workbook encryption and recommend stronger protection (C#)
// Description: Loads an Excel file with Aspose.Cells, inspects the Workbook.Settings.IsDefaultEncrypted flag, writes a console message about the encryption status, and advises using SetEncryptionOptions for a more robust algorithm before disposing the workbook.
// Keywords: Aspose.Cells default encryption detection | Workbook.IsDefaultEncrypted C# | upgrade Excel encryption Aspose | SetEncryptionOptions example | Excel file security compliance .NET
// Common Searches: How to know if an Excel workbook uses Aspose.Cells default encryption | C# code to verify encryption strength of a workbook with Aspose.Cells | Suggest stronger encryption for encrypted Excel files in .NET | Detect default encryption algorithm in Excel using Aspose.Cells
// Developer Intent: Identify whether a loaded workbook is protected with the library’s built‑in encryption and output guidance to apply a more secure algorithm.
// Use Cases: Audit incoming Excel documents for compliance by flagging those that rely on the built‑in encryption. | Run a scheduled scan of a file repository, log any workbooks using the default scheme, and generate a security report. | Embed the check in an upload service to warn users and prompt them to re‑encrypt with a stronger algorithm.
// AI Prompts: Generate C# code that re‑encrypts a workbook with AES‑256 using Aspose.Cells SetEncryptionOptions after detecting IsDefaultEncrypted. | Show how to configure Aspose.Cells to save a workbook with a custom password and a high‑strength encryption method. | Explain exception handling for loading an Excel file encrypted with the default algorithm in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // Loads an Excel file with Aspose.Cells, inspects the Workbook.Settings.IsDefaultEncrypted flag, writes a console message about the encryption status, and advises using SetEncryptionOptions for a more robust algorithm before disposing the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be inspected
            string workbookPath = "input.xlsx";

            // Load the workbook (create LoadOptions if needed)
            Workbook workbook = new Workbook(workbookPath);

            // Check if the workbook is encrypted with the default encryption algorithm
            bool isDefaultEncrypted = workbook.Settings.IsDefaultEncrypted;

            // Log the result and recommendation
            if (isDefaultEncrypted)
            {
                Console.WriteLine("The workbook uses the default encryption algorithm.");
                Console.WriteLine("Recommendation: Upgrade to a stronger encryption algorithm using SetEncryptionOptions.");
            }
            else
            {
                Console.WriteLine("The workbook does not use the default encryption algorithm.");
            }

            // Optional: Dispose the workbook if no further processing is required
            workbook.Dispose();
        }
    }
}

// Title: Save a Digitally Signed Excel Workbook as a New Copy with Aspose.Cells for .NET
// Description: Loads a signed Excel file, verifies the digital signature via Workbook.IsDigitallySigned, and saves it to a different filename using SaveFormat.Xlsx. The signature is retained and the source workbook stays untouched, enabling an audit‑trail copy.
// Keywords: Aspose.Cells | C# | .NET | digital signature | signed workbook copy | Workbook.IsDigitallySigned | SaveFormat.Xlsx | preserve signature | audit trail Excel | copy Excel file
// Common Searches: Aspose.Cells save signed workbook as copy | C# preserve digital signature when copying Excel file | how to create audit trail for signed Excel workbook .NET | check Workbook.IsDigitallySigned before saving | duplicate signed Excel file using Aspose.Cells
// Developer Intent: Create a duplicate of a digitally signed workbook while keeping the original unchanged and retaining its signature.
// Use Cases: Archive a signed financial statement by saving a timestamped copy without altering the source file. | Generate versioned backups of contract workbooks that require a verifiable digital signature. | Integrate signature verification and copy creation into a document‑management system that mandates immutable originals.
// AI Prompts: Generate C# code with Aspose.Cells that loads a signed workbook, confirms its digital signature, and saves it under a new name preserving the signature. | Show how to add comprehensive error handling for missing files and save failures when copying a signed Excel workbook in .NET. | Explain the role of Workbook.IsDigitallySigned and why SaveFormat.Xlsx maintains the digital signature in the copied file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a signed Excel file, verifies the digital signature via Workbook.IsDigitallySigned, and saves it to a different filename using SaveFormat.Xlsx. The signature is retained and the source workbook stays untouched, enabling an audit‑trail copy.
    public class SaveSignedWorkbookCopy
    {
        public static void Run()
        {
            // Path to the original signed workbook
            string originalPath = "SignedWorkbook.xlsx";

            // Verify that the source file exists
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Error: The file \"{originalPath}\" was not found.");
                return;
            }

            try
            {
                // Load the signed workbook
                Workbook workbook = new Workbook(originalPath);

                // Verify that the workbook is digitally signed
                bool isSigned = workbook.IsDigitallySigned;
                Console.WriteLine($"Original workbook is digitally signed: {isSigned}");

                // Define a new filename to preserve the original file (audit trail)
                string copyPath = "SignedWorkbook_Copy.xlsx";

                // Save the workbook to the new file (signature is preserved)
                workbook.Save(copyPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as a copy to: {copyPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SaveSignedWorkbookCopy.Run();
        }
    }
}

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SaveSignedWorkbookDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the original signed workbook
            string originalPath = "SignedOriginal.xlsx";

            // Verify the source file exists to avoid FileNotFoundException
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Source file not found: {originalPath}");
                return;
            }

            try
            {
                // Load the signed workbook
                Workbook workbook = new Workbook(originalPath);

                // Optional: verify that the workbook is digitally signed
                if (workbook.IsDigitallySigned)
                {
                    Console.WriteLine("Original workbook is digitally signed.");
                }
                else
                {
                    Console.WriteLine("Original workbook is NOT digitally signed.");
                }

                // Define the new file name for the audit copy
                string auditCopyPath = "SignedAuditCopy.xlsx";

                // Save the workbook under the new name (preserves the signature)
                workbook.Save(auditCopyPath);
                Console.WriteLine($"Workbook saved as audit copy: {auditCopyPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the error
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
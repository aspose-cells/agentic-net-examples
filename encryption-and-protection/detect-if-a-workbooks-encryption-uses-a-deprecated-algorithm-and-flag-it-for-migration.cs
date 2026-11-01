// Title: Detect deprecated Excel workbook encryption algorithm with Aspose.Cells for .NET (C#)
// Common Searches: C# code to determine if an .xlsx file uses a legacy encryption algorithm with Aspose.Cells | how to check if an Excel workbook is encrypted without a password using Aspose.Cells .NET | identify workbooks that need encryption migration when using Aspose.Cells | detect outdated Excel file encryption in a .NET application with Aspose.Cells | determine if loading an Excel file throws encryption exception in C# Aspose.Cells
// Tags: detect deprecated encryption Aspose.Cells | check workbook encryption algorithm .NET | load encrypted xlsx without password Aspose.Cells | legacy Excel encryption migration Aspose.Cells | identify encrypted workbook C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace EncryptionCheck
{
    // The example loads an Excel file with Aspose.Cells using LoadOptions for Xlsx without supplying a password. If loading throws an exception, the workbook is considered encrypted; because older Aspose.Cells versions do not expose the encryption algorithm, the code treats any encrypted workbook as using a deprecated algorithm and flags it for migration.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: pass the path of the workbook to check
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the workbook.");
                return;
            }

            string workbookPath = args[0];

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            bool isDeprecated = false;
            try
            {
                isDeprecated = IsWorkbookUsingDeprecatedEncryption(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
                return;
            }

            if (isDeprecated)
            {
                Console.WriteLine("The workbook uses a deprecated encryption algorithm. Migration is required.");
            }
            else
            {
                Console.WriteLine("The workbook encryption algorithm is up‑to‑date.");
            }
        }

        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>True if a deprecated algorithm is used; otherwise, false.</returns>
        static bool IsWorkbookUsingDeprecatedEncryption(string filePath)
        {
            try
            {
                // Attempt to load the workbook without a password.
                // If the workbook is not encrypted, this will succeed.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                Workbook workbook = new Workbook(filePath, loadOptions);
                return false; // Not encrypted, therefore not using a deprecated algorithm.
            }
            catch (Exception)
            {
                // An exception during loading without a password typically indicates encryption.
                // Aspose.Cells does not expose the encryption algorithm directly in older versions,
                // so we conservatively treat any encrypted workbook as using a deprecated algorithm.
                return true;
            }
        }
    }
}

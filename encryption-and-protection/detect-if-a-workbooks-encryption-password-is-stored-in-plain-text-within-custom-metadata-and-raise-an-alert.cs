// Title: Detect Plain‑Text Passwords in Excel Custom Document Properties with AspNet Aspose.Cells
// Description: A .NET utility that loads an Excel file, checks if the workbook is encrypted, reads only the document‑properties metadata, scans custom properties for names containing "password", and raises an alert when a non‑empty string value is found. It also reports the workbook's encryption status.
// Keywords: Aspose.Cells | .NET | Excel password detection | custom document properties | metadata security | plain text password | workbook encryption check | document properties API | security audit | Excel protection
// Common Searches: Aspose.Cells detect password in custom metadata | check Excel workbook for plain text password | read custom document properties with Aspose.Cells | Excel file security audit .NET | find exposed password in workbook metadata
// Developer Intent: Scan an Excel workbook's custom document properties for any plain‑text password entries and generate an alert if such a value is discovered.
// Use Cases: Automated security scan of a repository of Excel files to ensure passwords are not stored in custom metadata before release. | CI/CD gate that fails the build when a workbook contains a plain‑text password in its custom properties. | Runtime validation of user‑uploaded Excel files, warning users if a password is exposed in the file's metadata.
// AI Prompts: Create a reusable Aspose.Cells method that returns the names of custom document properties containing the word "password" with non‑empty values. | Write unit tests for the plain‑text password detection logic covering encrypted, unencrypted, and non‑password property scenarios. | Suggest a metadata‑only approach to detect exposed passwords without loading the full workbook, using Aspose.Cells APIs.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsSecurityCheck
{
    // A .NET utility that loads an Excel file, checks if the workbook is encrypted, reads only the document‑properties metadata, scans custom properties for names containing "password", and raises an alert when a non‑empty string value is found. It also reports the workbook's encryption status.
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be inspected
            string filePath = "sample.xlsx";

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' not found.");
                return;
            }

            try
            {
                // Load the workbook to check if it is encrypted (standard workbook loading)
                Workbook workbook = new Workbook(filePath);
                bool isWorkbookEncrypted = workbook.Settings.IsEncrypted;

                // Load only the document properties metadata (no need to load the whole workbook again)
                MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
                WorkbookMetadata metadata = new WorkbookMetadata(filePath, metaOptions);

                // Iterate through all custom document properties
                foreach (var prop in metadata.CustomDocumentProperties)
                {
                    // Check if the property name suggests it holds a password
                    if (prop.Name.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // If the value is a non‑empty string, it is stored in plain text
                        if (prop.Value is string plainText && !string.IsNullOrWhiteSpace(plainText))
                        {
                            Console.WriteLine($"ALERT: Encryption password stored in plain text in custom metadata property '{prop.Name}'.");
                        }
                    }
                }

                // Additional check: if the workbook itself is encrypted, the password should not be in plain text
                if (isWorkbookEncrypted)
                {
                    Console.WriteLine("Workbook is encrypted. Ensure the password is not exposed in metadata.");
                }
                else
                {
                    Console.WriteLine("Workbook is not encrypted.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}

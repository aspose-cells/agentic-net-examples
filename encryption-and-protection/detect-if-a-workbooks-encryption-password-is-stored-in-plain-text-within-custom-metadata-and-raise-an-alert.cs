// Title: Detect Plain‑Text Workbook Encryption Password in Custom Document Properties with Aspose.Cells for .NET
// Description: Loads an Excel file using Aspose.Cells, reads the workbook's encryption password via Workbook.Settings.Password, scans only the custom document properties, and raises an alert when any property value matches the password, indicating a plain‑text exposure.
// Keywords: Aspose.Cells password detection | Excel encryption password leak | custom document properties | plain text password check | C# workbook security | metadata inspection Aspose | .NET Excel protection audit
// Common Searches: How to find if an Excel password is stored in custom properties with Aspose.Cells | Aspose.Cells .NET check for plain‑text workbook password in metadata | Detect password leakage in Excel document properties using C# | Scan Excel file for exposed encryption password via Aspose.Cells
// Developer Intent: Determine whether the workbook’s encryption password appears as plain text in any custom document property and generate an alert if it does.
// Use Cases: Security audit of generated Excel reports to ensure passwords are not embedded in metadata. | Automated CI/CD gate that fails a build when a password is found in custom properties. | Compliance reporting that logs instances of password exposure for governance reviews.
// AI Prompts: Generate C# code with Aspose.Cells that iterates over all custom document properties of a workbook and flags any property whose value equals the workbook’s encryption password. | Create a method that returns the names of custom properties containing the workbook password in clear text. | Explain how to modify the solution to handle workbooks encrypted with an unknown password while still checking metadata for password exposure.

using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace WorkbookPasswordMetadataCheck
{
    // Loads an Excel file using Aspose.Cells, reads the workbook's encryption password via Workbook.Settings.Password, scans only the custom document properties, and raises an alert when any property value matches the password, indicating a plain‑text exposure.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be inspected
            string filePath = "input.xlsx";

            // Load the workbook (if the workbook is encrypted, the password must be supplied here)
            // For the purpose of this check we assume the workbook can be opened without a password
            // or that the correct password is already known.
            Workbook workbook = new Workbook(filePath);

            // Retrieve the encryption password set on the workbook (if any)
            string workbookPassword = workbook.Settings.Password;

            // If there is no password, there is nothing to check
            if (string.IsNullOrEmpty(workbookPassword))
            {
                Console.WriteLine("The workbook is not encrypted or no password is set.");
                return;
            }

            // Load only the document properties metadata (custom properties are part of this)
            MetadataOptions metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
            WorkbookMetadata metadata = new WorkbookMetadata(filePath, metaOptions);

            // Iterate through all custom document properties and compare their values with the password
            foreach (var customProp in metadata.CustomDocumentProperties)
            {
                // Ensure the property value is a string before comparison
                if (customProp.Value is string propValue && propValue == workbookPassword)
                {
                    // Alert: password found in plain text within custom metadata
                    Console.WriteLine($"Alert: Encryption password stored in plain text in custom metadata property '{customProp.Name}'.");
                }
            }

            // Optional: indicate completion if no issues were found
            Console.WriteLine("Metadata inspection completed.");
        }
    }
}

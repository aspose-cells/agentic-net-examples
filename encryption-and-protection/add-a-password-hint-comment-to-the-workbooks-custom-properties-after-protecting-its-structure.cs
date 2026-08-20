// Title: Add a Password Hint via Custom Document Property After Protecting Workbook Structure (Aspose.Cells for .NET)
// Description: Demonstrates how to protect an Excel workbook's structure with a password using Aspose.Cells, then store a user‑friendly password hint in a custom document property, save the file, and optionally read the hint back for verification.
// Keywords: Aspose.Cells protect workbook structure | Excel password hint custom property | Aspose.Cells add custom document property | C# store password hint in Excel | Aspose.Cells .NET encryption example
// Common Searches: how to add a password hint to an Excel file with Aspose.Cells | protect workbook structure and store hint Aspose.Cells C# | read custom document property from a protected workbook | Aspose.Cells example for structure protection and hint
// Developer Intent: Create a protected workbook and embed a readable password hint as a custom document property.
// Use Cases: Provide administrators with a non‑intrusive hint for a protected workbook password. | Display the hint in a UI before prompting users for the protection password. | Update or replace the hint later without removing the existing structure protection.
// AI Prompts: Generate C# code with Aspose.Cells that protects a workbook's structure and adds a custom document property named "PasswordHint" containing a hint. | Show how to retrieve the "PasswordHint" property from a workbook that has structure protection applied using Aspose.Cells for .NET. | Explain how to modify the password hint in an already protected workbook without disabling the structure protection.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordHintDemo
{
    // Demonstrates how to protect an Excel workbook's structure with a password using Aspose.Cells, then store a user‑friendly password hint in a custom document property, save the file, and optionally read the hint back for verification.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a password
            // This prevents adding, deleting, renaming, moving, or hiding worksheets without the password
            workbook.Protect(ProtectionType.Structure, "MyStrongPassword!");

            // Add a custom document property that serves as a password hint
            // Custom properties are stored in the workbook and can be read without opening the file
            workbook.CustomDocumentProperties.Add("PasswordHint", "Your favorite pet's name");

            // Save the workbook to a file
            workbook.Save("Workbook_With_PasswordHint.xlsx");

            // Optional: Verify that the custom property was added
            Console.WriteLine("Custom Property 'PasswordHint' = " +
                workbook.CustomDocumentProperties["PasswordHint"].Value);
        }
    }
}

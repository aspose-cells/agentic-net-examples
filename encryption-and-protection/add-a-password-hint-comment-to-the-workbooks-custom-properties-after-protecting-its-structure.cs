// Title: Add a Password Hint via Custom Document Property After Protecting Workbook Structure with Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, apply structure protection with a password, embed a readable password hint as a custom document property, and save the file as an XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom document property | Excel workbook structure protection | password hint Excel file | C# protect workbook Aspose.Cells | store hint in Excel custom property
// Common Searches: Aspose.Cells add password hint after protecting workbook | How to store a password hint in Excel using custom properties | Protect workbook structure and include hint with Aspose.Cells .NET | Read password hint from a protected Excel file
// Developer Intent: Insert a readable password hint into a workbook’s custom document properties after applying structure protection.
// Use Cases: Generate template workbooks that are structure‑locked yet provide users with a hint for the opening password. | Automate creation of multiple protected spreadsheets, each containing a unique hint for its password. | Allow administrators to embed recovery clues without compromising workbook protection.
// AI Prompts: Write C# code with Aspose.Cells that protects a workbook’s structure and adds a custom document property named "PasswordHint" containing a user‑defined clue. | Show how to retrieve the "PasswordHint" custom property from a structure‑protected Excel file using Aspose.Cells. | Explain how to modify the password hint in an already protected workbook without removing the existing protection.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordHintDemo
{
    // Demonstrates how to create a new Workbook, apply structure protection with a password, embed a readable password hint as a custom document property, and save the file as an XLSX using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a password
            // This prevents adding, deleting, renaming, moving, hiding, or unhiding worksheets
            workbook.Protect(ProtectionType.Structure, "StrongPassword!123");

            // Add a custom document property that serves as a password hint
            // Custom properties are stored in the workbook and can be viewed without opening the file
            workbook.CustomDocumentProperties.Add("PasswordHint", "First pet's name");

            // Save the workbook to a file
            workbook.Save("Workbook_With_PasswordHint.xlsx", SaveFormat.Xlsx);
        }
    }
}

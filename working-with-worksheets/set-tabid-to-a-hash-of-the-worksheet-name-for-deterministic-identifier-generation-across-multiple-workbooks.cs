// Title: Create deterministic Worksheet TabId values by hashing sheet names with Aspose.Cells for .NET
// AI Prompts: Write a C# method that iterates through all worksheets in a workbook, computes a stable 32‑bit integer hash of each sheet's Name using SHA‑256, and assigns the result to Worksheet.TabId via Aspose.Cells. | Generate code that loads an Excel file, applies a hash‑based TabId to every worksheet for consistent identification across workbooks, and saves the updated file.
// Common Searches: C# Aspose.Cells set TabId from worksheet name hash | deterministic TabId generation for Excel sheets using SHA256 | how to assign stable integer IDs to worksheets in Aspose.Cells | consistent worksheet identifiers across multiple workbooks .NET | hash sheet name to 32‑bit integer for TabId property
// Tags: set worksheet TabId using hash Aspose.Cells | deterministic TabId generation C# | SHA256 hash for worksheet identifier | hash sheet name to int32 | consistent TabId across workbooks

using System;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The example loads a workbook, iterates through each worksheet, computes a 32‑bit integer hash of the worksheet name using SHA‑256, assigns this value to the Worksheet.TabId property, and saves the modified workbook, ensuring the same TabId is reproduced for identical sheet names across different files.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and assign a deterministic TabId
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Compute a stable hash from the worksheet name
            int tabId = ComputeDeterministicHash(sheet.Name);
            // Set the TabId property
            sheet.TabId = tabId;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }

    // Generates a deterministic 32‑bit integer hash from a string using SHA‑256
    static int ComputeDeterministicHash(string input)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] data = Encoding.UTF8.GetBytes(input);
            byte[] hash = sha256.ComputeHash(data);
            // Use the first 4 bytes of the hash to form an int (little‑endian)
            return BitConverter.ToInt32(hash, 0);
        }
    }
}

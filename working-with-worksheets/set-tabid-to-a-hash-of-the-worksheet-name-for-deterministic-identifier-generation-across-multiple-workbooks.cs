// Title: Set a deterministic TabId for each worksheet by hashing the sheet name with SHA‑256 using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through all worksheets in a workbook, computes a 32‑bit SHA‑256 hash of each sheet's Name, and assigns the result to the Worksheet.TabId property with Aspose.Cells. | Create a reusable method that converts a string to an int via the first four bytes of a SHA‑256 hash and demonstrate its use to give stable TabId values across multiple Excel files.
// Common Searches: how to generate stable TabId values for Excel worksheets in Aspose.Cells C# | C# compute 32‑bit hash from worksheet name for consistent sheet identifiers | using SHA256 to create deterministic worksheet IDs with Aspose.Cells .NET | assign integer TabId to each sheet based on its name in a .xlsx file | deterministic worksheet identifiers across multiple workbooks Aspose.Cells example
// Tags: deterministic worksheet TabId assignment Aspose.Cells | SHA256 string to int conversion C# | stable Excel sheet identifiers .NET | hash based TabId generation for worksheets | worksheet name hashing for consistent IDs

using System;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The program loads an Excel workbook, computes a 32‑bit SHA‑256 hash of each worksheet's name, assigns that hash to the worksheet's TabId property, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and set a deterministic TabId
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Compute a stable 32‑bit hash from the worksheet name
            int hash = ComputeDeterministicHash(sheet.Name);

            // Assign the hash to the worksheet's TabId (assumes TabId property exists)
            sheet.TabId = hash;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Generates a 32‑bit integer hash from a string using SHA256 (first 4 bytes)
    private static int ComputeDeterministicHash(string input)
    {
        using (SHA256 sha = SHA256.Create())
        {
            byte[] data = Encoding.UTF8.GetBytes(input);
            byte[] hash = sha.ComputeHash(data);
            // Convert first 4 bytes to an int (little‑endian)
            return BitConverter.ToInt32(hash, 0);
        }
    }
}

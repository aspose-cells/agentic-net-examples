// Title: C# – Generate SHA256 checksum from all cell values in an Aspose.Cells workbook
// Description: The sample creates a Workbook, fills cells with text, numbers, dates and booleans, walks through every cell via Cells.GetEnumerator, appends each formatted value to a StringBuilder, hashes the combined text with SHA256, converts the digest to hex, prints it and saves the file as ChecksumDemo.xlsx.
// Keywords: Aspose.Cells | C# checksum | SHA256 hash Excel | cell enumeration | string concatenation | .NET data integrity | Workbook hash | Excel verification | hashing cell values
// Common Searches: how to compute SHA256 hash of Excel cells using Aspose.Cells C# | enumerate cells in Aspose.Cells and create a checksum | concatenate cell values and generate a hash in .NET | verify workbook integrity with SHA256 in Aspose.Cells | C# example for hashing all worksheet values
// Developer Intent: Create a SHA256 digest that represents the combined string values of every populated cell in a worksheet.
// Use Cases: Validate that a workbook has not been altered during transfer by comparing hashes before and after upload. | Detect accidental or malicious changes to worksheet content by recomputing the digest. | Generate a stable identifier for worksheet data to support caching, version control, or duplicate detection.
// AI Prompts: Write a reusable method that returns the SHA256 hash of all non‑empty cells in a given Aspose.Cells worksheet. | Adapt the code to ignore formula cells and hash only the displayed results. | Show how to store the computed hash in a hidden worksheet cell and retrieve it later for verification.

using System;
using System.Collections;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

// The sample creates a Workbook, fills cells with text, numbers, dates and booleans, walks through every cell via Cells.GetEnumerator, appends each formatted value to a StringBuilder, hashes the combined text with SHA256, converts the digest to hex, prints it and saves the file as ChecksumDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue(123);
        worksheet.Cells["C1"].PutValue(DateTime.Now);
        worksheet.Cells["A2"].PutValue(3.14);
        worksheet.Cells["B2"].PutValue(true);

        // Enumerate all cells using the Cells.GetEnumerator method
        StringBuilder concatenatedValues = new StringBuilder();
        IEnumerator enumerator = worksheet.Cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            if (cell != null && cell.Value != null)
            {
                // Use the formatted string representation of the cell value
                concatenatedValues.Append(cell.StringValue);
            }
        }

        // Compute a SHA256 checksum of the concatenated string
        byte[] hashBytes;
        using (SHA256 sha256 = SHA256.Create())
        {
            hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(concatenatedValues.ToString()));
        }

        // Convert the hash to a hexadecimal string for display
        StringBuilder hex = new StringBuilder(hashBytes.Length * 2);
        foreach (byte b in hashBytes)
        {
            hex.AppendFormat("{0:x2}", b);
        }

        Console.WriteLine("Checksum (SHA256): " + hex.ToString());

        // Save the workbook (save rule)
        workbook.Save("ChecksumDemo.xlsx");
    }
}

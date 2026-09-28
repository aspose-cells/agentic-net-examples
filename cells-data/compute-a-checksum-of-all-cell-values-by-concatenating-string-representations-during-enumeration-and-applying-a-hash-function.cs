// Title: Generate a SHA256 checksum of all cell values in an Aspose.Cells workbook with C#
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through every cell in a worksheet, concatenate each non‑null value as a string, and return the SHA256 hash as a hex string. | Create a reusable method in C# that accepts an Aspose.Cells Workbook and produces a checksum by hashing the concatenated string representations of all cell values. | Modify the example to compute an MD5 checksum instead of SHA256 while keeping the same cell enumeration and concatenation logic.
// Common Searches: C# Aspose.Cells how to hash all worksheet cell values | calculate SHA256 checksum for Excel data using Aspose.Cells | enumerate cells in Aspose.Cells and generate a data checksum | concatenate non‑null cell values and compute hash in .NET | verify workbook integrity with checksum in Aspose.Cells C#
// Tags: Aspose.Cells compute SHA256 checksum | enumerate worksheet cells C# | concatenate cell values Aspose.Cells | hash workbook data .NET | cell value checksum Excel

using System;
using System.Collections;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsChecksumDemo
{
    // The program creates a workbook, fills several cells with different data types, enumerates all cells, concatenates their non‑null string representations, computes a SHA256 hash of the combined string, outputs the hexadecimal checksum, and saves the workbook as ChecksumDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Hello");
            cells["B1"].PutValue(123);
            cells["C1"].PutValue(DateTime.Now);
            cells["A2"].PutValue(3.1415);
            cells["B2"].PutValue(true);
            cells["C2"].PutValue("World");

            // Concatenate string representations of all cell values
            StringBuilder concatenated = new StringBuilder();

            // Enumerate cells using the provided GetEnumerator method
            IEnumerator enumerator = cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                // Use the cell's Value property; if null, skip
                if (cell.Value != null)
                {
                    concatenated.Append(cell.Value.ToString());
                }
            }

            // Compute a SHA256 hash of the concatenated string
            byte[] hashBytes;
            using (SHA256 sha = SHA256.Create())
            {
                hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(concatenated.ToString()));
            }

            // Convert hash bytes to a hexadecimal string
            StringBuilder hashString = new StringBuilder();
            foreach (byte b in hashBytes)
            {
                hashString.Append(b.ToString("x2"));
            }

            // Output the checksum
            Console.WriteLine("Checksum (SHA256) of all cell values:");
            Console.WriteLine(hashString.ToString());

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChecksumDemo.xlsx");
        }
    }
}

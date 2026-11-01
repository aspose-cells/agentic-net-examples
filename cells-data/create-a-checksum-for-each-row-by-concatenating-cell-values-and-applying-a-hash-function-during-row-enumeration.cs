// Title: Add a SHA‑256 checksum column to each data row in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, concatenates the values of every cell in each data row, computes a SHA‑256 hash, and writes the hex string into a new column at the end of the sheet. | Create a .NET console application that iterates through rows of a workbook, builds a single string from all cell values per row, uses System.Security.Cryptography.SHA256 to produce a checksum, and saves the workbook with the checksum column. | Modify an existing Aspose.Cells workbook to insert a 'Checksum' header after the last data column and populate each row with the SHA‑256 hash of its concatenated cell contents.
// Common Searches: how to compute a SHA256 hash for each row in an Excel file using Aspose.Cells C# | Aspose.Cells add checksum column to worksheet for data integrity | C# example to concatenate row values and generate SHA256 checksum in Excel | save row-level hash in new column with Aspose.Cells .NET | enumerate rows and calculate SHA256 checksum per row in .xlsx using Aspose.Cells
// Tags: SHA256 checksum per Excel row Aspose.Cells | append hash column to worksheet .NET | concatenate cell values for row hashing | row-level data integrity verification Aspose.Cells | C# compute SHA256 for Excel rows | use System.Security.Cryptography with Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

// The example loads an Excel workbook, determines the used range, adds a 'Checksum' header after the last data column, iterates each data row, concatenates all cell values, computes a SHA‑256 hash of the concatenated string, writes the hexadecimal hash into the new column, and saves the updated file.
class RowChecksumGenerator
{
    static void Main(string[] args)
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output_with_checksum.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the used range
        int maxRow = cells.MaxDataRow;      // zero‑based index of the last row with data
        int maxCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

        // Define the column where the checksum will be stored (one column after the last data column)
        int checksumColIndex = maxCol + 1;
        cells[0, checksumColIndex].PutValue("Checksum"); // header

        // Prepare SHA256 hash algorithm
        using (SHA256 sha256 = SHA256.Create())
        {
            // Enumerate each data row (skip header row if present)
            for (int row = 1; row <= maxRow; row++)
            {
                StringBuilder sb = new StringBuilder();

                // Concatenate all cell values in the current row
                for (int col = 0; col <= maxCol; col++)
                {
                    object value = cells[row, col].Value;
                    if (value != null)
                    {
                        sb.Append(value.ToString());
                    }
                    // Optional delimiter can be added, e.g., sb.Append("|");
                }

                // Compute hash of the concatenated string
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));

                // Convert hash to hexadecimal string
                StringBuilder hashString = new StringBuilder(hashBytes.Length * 2);
                foreach (byte b in hashBytes)
                {
                    hashString.Append(b.ToString("x2"));
                }

                // Write checksum into the designated column
                cells[row, checksumColIndex].PutValue(hashString.ToString());
            }
        }

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save(outputPath);
    }
}

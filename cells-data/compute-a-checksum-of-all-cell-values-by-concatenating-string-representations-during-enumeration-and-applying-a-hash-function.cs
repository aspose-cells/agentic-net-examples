// Title: Compute a SHA-256 checksum of all cell values in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells generate SHA256 checksum from worksheet cell values | How to create a data integrity hash for an Excel file with Aspose.Cells | Iterate used range and hash concatenated cell data in .NET | Calculate workbook checksum by concatenating cell values in C# | Aspose.Cells compute SHA-256 hash of spreadsheet contents
// Tags: compute SHA-256 hash from Aspose.Cells used range | concatenate worksheet cell values for checksum in C# | generate Excel workbook checksum with Aspose.Cells | data integrity hash for Excel using .NET | Aspose.Cells iterate cells for SHA256 calculation

using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet's used range, concatenates the string representation of every cell value, converts the combined string to UTF‑8 bytes, computes a SHA‑256 hash, and returns the hash as a hexadecimal string, while handling missing files and runtime errors.
public class WorkbookChecksum
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>Hexadecimal string of the SHA‑256 hash.</returns>
    public static string ComputeChecksum(string filePath)
    {
        if (!File.Exists(filePath))
        {
            Console.WriteLine("File not found: " + filePath);
            return string.Empty;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(filePath);

            // StringBuilder to accumulate cell values
            StringBuilder sb = new StringBuilder();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet to avoid iterating empty cells
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // If the worksheet is empty, skip it
                if (usedRange == null)
                    continue;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop rows
                for (int row = startRow; row <= endRow; row++)
                {
                    // Loop columns
                    for (int col = startCol; col <= endCol; col++)
                    {
                        // Retrieve the cell
                        Cell cell = sheet.Cells[row, col];

                        // Append its string representation (empty string for null)
                        sb.Append(cell.Value?.ToString() ?? string.Empty);
                    }
                }
            }

            // Convert the concatenated string to bytes (UTF‑8)
            byte[] dataBytes = Encoding.UTF8.GetBytes(sb.ToString());

            // Compute SHA‑256 hash
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(dataBytes);

                // Convert hash bytes to hexadecimal string
                StringBuilder hashSb = new StringBuilder(hashBytes.Length * 2);
                foreach (byte b in hashBytes)
                    hashSb.Append(b.ToString("x2"));

                return hashSb.ToString();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error computing checksum: {ex.Message}");
            return string.Empty;
        }
    }

    // Example usage
    public static void Main()
    {
        string excelPath = @"C:\Path\To\Your\File.xlsx";

        if (!File.Exists(excelPath))
        {
            Console.WriteLine("File not found: " + excelPath);
            return;
        }

        try
        {
            string checksum = ComputeChecksum(excelPath);
            if (!string.IsNullOrEmpty(checksum))
                Console.WriteLine("Checksum (SHA-256): " + checksum);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

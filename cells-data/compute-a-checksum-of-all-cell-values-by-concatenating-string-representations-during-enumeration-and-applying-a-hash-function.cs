// Title: Generate a SHA-256 checksum of all cell values in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to create a SHA-256 hash from all cells in an Excel workbook | Compute checksum of Excel data using Aspose.Cells .NET library | Iterate used range and hash concatenated cell values with SHA256 in C# | Generate workbook integrity hash by concatenating cell values with Aspose.Cells | Calculate SHA256 checksum for Excel file content using Aspose.Cells API
// Tags: SHA-256 hash generation from Aspose.Cells workbook | enumerate used range cells Aspose.Cells | concatenate cell values for checksum Aspose.Cells | C# compute Excel workbook checksum | Aspose.Cells data integrity verification

using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, iterates each worksheet's used range, concatenates the string representation of every cell value (treating nulls as empty), converts the combined string to UTF-8 bytes, computes a SHA-256 hash, and returns the hash as a hexadecimal string for data‑integrity verification.
public class ChecksumCalculator
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>Hexadecimal string of the SHA‑256 hash.</returns>
    public static string ComputeChecksum(string filePath)
    {
        // Verify the file exists before attempting to load
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel file not found.", filePath);

        Workbook workbook;
        try
        {
            // Load the workbook using Aspose.Cells
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to load workbook.", ex);
        }

        // StringBuilder to accumulate cell values
        StringBuilder sb = new StringBuilder();

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range of the worksheet to limit enumeration
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            // Determine start and end indices
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Loop through rows and columns within the used range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    // Retrieve the cell
                    Cell cell = sheet.Cells[row, col];

                    // Append the string representation of the cell's value
                    // Null or empty cells are represented as an empty string
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

            // Convert hash bytes to a hexadecimal string
            StringBuilder hashSb = new StringBuilder(hashBytes.Length * 2);
            foreach (byte b in hashBytes)
                hashSb.Append(b.ToString("x2"));

            return hashSb.ToString();
        }
    }

    // Example usage
    public static void Main()
    {
        string excelPath = @"C:\Path\To\Your\Workbook.xlsx";

        if (!File.Exists(excelPath))
        {
            Console.WriteLine("File not found: " + excelPath);
            return;
        }

        try
        {
            string checksum = ComputeChecksum(excelPath);
            Console.WriteLine("Checksum (SHA-256): " + checksum);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while computing the checksum: " + ex.Message);
        }
    }
}

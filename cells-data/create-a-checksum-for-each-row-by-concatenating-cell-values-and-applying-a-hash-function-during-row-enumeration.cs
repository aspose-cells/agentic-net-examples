// Title: Generate a SHA256 checksum for each Excel row by concatenating cell values with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add a checksum column to each row of an .xlsx workbook | How to compute a SHA256 hash of concatenated row values using Aspose.Cells .NET | Create per‑row hash in Excel with Aspose.Cells and save to a new column | Aspose.Cells example for generating row‑wise checksum in a spreadsheet | Append SHA256 checksum to every row in an existing Excel file using C#
// Tags: row checksum generation Aspose.Cells | SHA256 hash Excel rows C# | add checksum column worksheet Aspose.Cells | concatenate row values Aspose.Cells | per‑row hash .xlsx C# | Excel data integrity verification Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using Aspose.Cells;

// The program loads an existing workbook (or creates a sample one), determines the next empty column, iterates through each populated row, concatenates all cell values in that row, computes a SHA256 hash of the concatenated string, writes the hash into the new checksum column, and saves the modified workbook, with handling for password‑protected files.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook or create a new one if it does not exist
            if (File.Exists(inputPath))
            {
                // Load the workbook (no password assumed)
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with sample data
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample");
                sheet.Cells["B1"].PutValue(123);
                sheet.Cells["A2"].PutValue("Test");
                sheet.Cells["B2"].PutValue(456);
                workbook.Save(inputPath);
            }

            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the column index where the checksum will be stored (next column after the last used)
            int checksumColumnIndex = cells.MaxColumn + 1; // zero‑based index

            // Iterate over each row that contains data
            for (int row = cells.MinRow; row <= cells.MaxRow; row++)
            {
                // Concatenate all cell values in the current row
                StringBuilder concatenated = new StringBuilder();
                for (int col = cells.MinColumn; col <= cells.MaxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell != null && cell.Value != null)
                    {
                        concatenated.Append(cell.Value.ToString());
                    }
                }

                // Compute SHA256 hash of the concatenated string
                string checksum = ComputeSha256Hash(concatenated.ToString());

                // Write the checksum into the designated column
                cells[row, checksumColumnIndex].PutValue(checksum);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
        }
        catch (CellsException ex) when (ex.Message.Contains("password"))
        {
            Console.WriteLine("The workbook is password protected. Please provide a valid password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Returns SHA256 hash of a string as a hexadecimal representation
    static string ComputeSha256Hash(string rawData)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] bytes = Encoding.UTF8.GetBytes(rawData);
            byte[] hashBytes = sha256.ComputeHash(bytes);
            StringBuilder builder = new StringBuilder(hashBytes.Length * 2);
            foreach (byte b in hashBytes)
                builder.Append(b.ToString("x2"));
            return builder.ToString();
        }
    }
}

// Title: Detect whether an OOXML (.xlsx) workbook is password‑protected without a password using Aspose.Cells for .NET
// Common Searches: how to check if an .xlsx file is encrypted with Aspose.Cells in C# without providing a password | detect password protection on an Excel workbook using LoadOptions Auto in Aspose.Cells | determine if a workbook is encrypted before opening it with Aspose.Cells for .NET | identify encrypted OOXML workbook format without a password using Aspose.Cells C# | use CellsException to find out if an Excel file requires a password in Aspose.Cells
// Tags: aspocells encrypted file detection | loadoptions auto format usage | cellsexception password required handling | c# write detection result file | xlsx format identification without password

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the presence of an .xlsx file, creates a LoadOptions instance with LoadFormat.Auto, and attempts to load the workbook without a password. If loading succeeds, it records that the file is not encrypted and notes the detected format; if a CellsException indicating a missing password is caught, it records that the workbook is encrypted and specifies the OOXML format. All outcomes are written to result.txt.
class Program
{
    static void Main()
    {
        // Path to the encrypted OOXML workbook
        string filePath = "encrypted.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            File.WriteAllText("result.txt", $"Error: File '{filePath}' not found.");
            return;
        }

        // Use LoadOptions with auto detection of the format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);

        try
        {
            // Attempt to load the workbook without providing a password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, the workbook is not encrypted
            string format = loadOptions.LoadFormat.ToString();
            File.WriteAllText("result.txt", $"Not encrypted. Format: {format}");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException when a password is required
            if (ex.Message != null && ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // The workbook is encrypted; format is OOXML (Xlsx)
                string format = LoadFormat.Xlsx.ToString();
                File.WriteAllText("result.txt", $"Encrypted workbook detected. Format: {format}");
            }
            else
            {
                // Other CellsException errors
                File.WriteAllText("result.txt", $"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Any other unexpected error
            File.WriteAllText("result.txt", $"Error: {ex.Message}");
        }
    }
}

// Title: Detect an encrypted Excel workbook and show a warning using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells check if an .xlsx file is password protected without providing a password | How to determine workbook encryption status with Aspose.Cells in a .NET application | Catch CellsException to identify encrypted Excel workbook using Aspose.Cells C# | Validate Excel file encryption before processing it with Aspose.Cells | Display a warning when loading an encrypted workbook in Aspose.Cells for .NET
// Tags: detect encrypted workbook Aspose.Cells | handle password protected Excel .NET | catch CellsException encryption detection | validate workbook encryption C# | Aspose.Cells load warning for encrypted file

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells. If a CellsException is thrown, the code assumes the workbook is encrypted or password‑protected and prints a warning; otherwise it confirms successful loading, outputs the worksheet count, and handles other errors.
class Program
{
    static void Main()
    {
        // Path to the workbook to be checked
        string filePath = "input.xlsx";

        // Ensure the file exists before attempting to load
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File \"{filePath}\" not found.");
            return;
        }

        Workbook workbook = null;
        bool isEncrypted = false;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(filePath);
        }
        catch (CellsException)
        {
            // The file is likely encrypted or password protected
            isEncrypted = true;
        }
        catch (Exception ex)
        {
            // Handle other unexpected errors
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        if (isEncrypted)
        {
            Console.WriteLine("Warning: The workbook is encrypted.");
        }
        else
        {
            Console.WriteLine("Workbook loaded successfully. It is not encrypted.");
            // Additional processing can be done here using the 'workbook' instance
            try
            {
                // Example: output the number of worksheets
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

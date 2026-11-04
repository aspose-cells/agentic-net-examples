// Title: Export an unencrypted Excel workbook to CSV with Aspose.Cells for .NET while handling password‑protected files
// Common Searches: c# aspnet check if Excel workbook is password protected before converting to csv using aspose.cells | how to safely export a workbook to csv with aspose.cells when the file may be encrypted | asp.net core load xlsx and save as csv only if not encrypted aspose.cells | detect encrypted excel file in c# and skip csv export using aspose.cells | asp.net c# export to csv without prompting for password aspose.cells
// Tags: export workbook to csv aspose.cells c# | detect encrypted excel workbook aspose.cells | load xlsx without password aspose.cells | save as csv without decryption aspose.cells | handle password protected workbook aspose.cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an XLSX file with Aspose.Cells, catches password‑related exceptions to determine if the workbook is encrypted, and if it loads successfully, saves the workbook as a CSV file while handling any I/O errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Ensure the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex) when (ex.Message.Contains("Password"))
        {
            // The workbook is encrypted and cannot be opened without a password
            Console.WriteLine("The workbook is encrypted and cannot be exported to CSV.");
            return;
        }
        catch (Exception ex)
        {
            // Other loading errors
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Export the workbook to CSV format
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook exported to CSV successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any errors during saving
            Console.WriteLine($"Error saving CSV: {ex.Message}");
        }
    }
}

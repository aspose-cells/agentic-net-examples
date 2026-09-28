// Title: Detect legacy .xls format and apply password protection with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, checks the file extension, sets workbook.Settings.Password, and saves the file keeping its original extension. | Create a reusable method that receives inputPath, outputPath, and password, validates the file, determines if it is binary (.xls) or OpenXML (.xlsx), and applies opening password accordingly. | Demonstrate exception handling while encrypting both legacy and modern Excel files using Aspose.Cells, ensuring the output file retains the same format.
// Common Searches: how to add opening password to a .xls file using Aspose.Cells C# | encrypt legacy Excel binary workbook with Aspose.Cells without converting to .xlsx | detect Excel file type and set password protection in Aspose.Cells for .NET | apply password protection to Excel files while preserving original format Aspose.Cells | C# Aspose.Cells encrypt both .xls and .xlsx files with same code
// Tags: Aspose.Cells opening password for binary XLS | C# detect Excel workbook format Aspose.Cells | password protect legacy Excel file without conversion | preserve original file extension when encrypting Excel Aspose.Cells | use workbook.Settings.Password to secure Excel workbook

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file (either .xls or .xlsx) with Aspose.Cells, verifies its existence, sets an opening password via workbook.Settings.Password, and saves the workbook, letting Aspose.Cells infer the format from the file extension so the original file type is preserved.
class ExcelEncryption
{
    static void Main()
    {
        // Input Excel file path (could be .xls or .xlsx)
        string inputPath = "input.xls";

        // Output encrypted Excel file path
        string outputPath = "encrypted.xls";

        // Password to apply
        string password = "Secret123";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply opening password protection (works for both legacy and modern formats)
            workbook.Settings.Password = password;

            // Save the encrypted workbook; format is inferred from the file extension
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook encrypted and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

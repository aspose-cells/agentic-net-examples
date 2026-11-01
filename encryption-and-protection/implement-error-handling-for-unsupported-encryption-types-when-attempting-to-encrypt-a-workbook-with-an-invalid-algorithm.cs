// Title: C# – Catch NotSupportedException for unsupported encryption algorithm while encrypting an Excel workbook with Aspose.Cells
// Common Searches: how to handle unsupported encryption algorithm exception when saving an encrypted Excel file with Aspose.Cells in C# | Aspose.Cells encrypt workbook with password and catch NotSupportedException for invalid algorithm | C# example of error handling for unsupported encryption types in Aspose.Cells workbook encryption | catch NotSupportedException during Aspose.Cells workbook.Save with encryption in .NET | validate encryption algorithm before saving encrypted Excel using Aspose.Cells C#
// Tags: Aspose.Cells workbook encryption NotSupportedException handling | C# encrypt Excel workbook with password error handling | unsupported encryption algorithm Aspose.Cells | save encrypted workbook with Aspose.Cells C# | Excel file encryption exception handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, optionally adds data, sets a password, ensures the output directory exists, and saves the file as an encrypted XLSX. The save operation is wrapped in a try‑catch block that specifically catches NotSupportedException for unsupported encryption algorithms, ArgumentException for argument issues, and a generic Exception for any other unexpected errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Optional: add some data
            workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

            // Set password for encryption
            workbook.Settings.Password = "MyPassword";

            // NOTE: The EncryptionAlgorithm property may not exist in older Aspose.Cells versions.
            // If needed, it can be set via SaveOptions in newer versions. Here we rely on the default algorithm.

            // Define output path
            string outputPath = "EncryptedWorkbook.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the encrypted workbook (save rule)
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (NotSupportedException ex)
        {
            // Handle unsupported encryption algorithms
            Console.WriteLine($"Error: {ex.Message}");
        }
        catch (ArgumentException ex)
        {
            // Handle other argument errors
            Console.WriteLine($"Error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

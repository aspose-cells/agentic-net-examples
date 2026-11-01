// Title: Encrypt multiple Excel workbooks in a folder with a shared password using Aspose.Cells for .NET
// Common Searches: C# code to batch encrypt all .xlsx files in a directory with Aspose.Cells | How to apply the same password to many Excel workbooks using Aspose.Cells .NET | Save encrypted Excel files to a separate output folder with Aspose.Cells | Automate workbook password protection for multiple files in C# | Aspose.Cells example for encrypting workbooks in bulk
// Tags: batch encrypt Excel workbooks Aspose.Cells | shared password workbook protection C# | save encrypted .xlsx to output directory | Workbook.Settings.Password Aspose.Cells example | process multiple Excel files Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The program iterates through every .xlsx file in a specified input folder, loads each workbook with Aspose.Cells, assigns a common password via the Workbook.Settings.Password property, and saves the encrypted workbook to a designated output folder, creating the output directory if it does not exist.
class EncryptWorkbooks
{
    static void Main()
    {
        // Directory containing the original workbooks
        string inputDirectory = @"C:\InputWorkbooks";

        // Directory where encrypted workbooks will be saved
        string outputDirectory = @"C:\EncryptedWorkbooks";

        // Shared password for all workbooks
        string password = "MySecretPassword";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Process each .xlsx file in the input directory
        foreach (string inputPath in Directory.GetFiles(inputDirectory, "*.xlsx"))
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Apply password protection (encryption)
            workbook.Settings.Password = password;

            // Determine the output file path
            string fileName = Path.GetFileName(inputPath);
            string outputPath = Path.Combine(outputDirectory, fileName);

            // Save the encrypted workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }

        Console.WriteLine("Encryption completed for all workbooks.");
    }
}

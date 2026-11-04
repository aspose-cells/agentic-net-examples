// Title: Encrypt all Excel .xlsx workbooks in a folder with a shared password using Aspose.Cells for .NET (C#)
// Common Searches: C# code to apply the same password to every .xlsx file in a directory with Aspose.Cells | How to batch encrypt Excel workbooks using Aspose.Cells .NET library | Programmatically protect multiple Excel files with a single password in C# | Aspose.Cells example for encrypting all workbooks in a folder | Save encrypted copies of Excel files to a different folder using Aspose.Cells
// Tags: batch workbook encryption Aspose.Cells C# | apply shared password to .xlsx files | encrypt Excel workbooks during save | process directory of Excel files Aspose | output encrypted workbooks to separate folder

using System;
using System.IO;
using Aspose.Cells;

// The sample iterates through every .xlsx file in a specified input directory, loads each workbook with Aspose.Cells, sets a common password via workbook.Settings.Password, and saves the encrypted workbook to an output directory, creating the folder if it does not exist.
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
        foreach (string filePath in Directory.GetFiles(inputDirectory, "*.xlsx"))
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(filePath);

            // Apply password protection (encryption)
            workbook.Settings.Password = password;

            // Determine the output file path
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputDirectory, fileName);

            // Save the encrypted workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }

        Console.WriteLine("Encryption completed for all workbooks.");
    }
}

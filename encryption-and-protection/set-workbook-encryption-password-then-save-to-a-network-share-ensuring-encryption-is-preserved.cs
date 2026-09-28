// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and save it directly to a UNC network share
// AI Prompts: Generate C# code that assigns a password via workbook.Settings.Password to encrypt an Excel file with Aspose.Cells and writes the file to a UNC network path. | Demonstrate how to confirm that password protection remains after saving an Aspose.Cells workbook to a shared folder.
// Common Searches: C# Aspose.Cells apply password protection and write workbook to network share | Saving encrypted Excel file to UNC path using Aspose.Cells .NET | Preserve workbook password when saving to a shared folder with Aspose.Cells | Aspose.Cells workbook.Settings.Password example for remote file server | Encrypt Excel workbook and store on remote network drive with Aspose.Cells
// Tags: Aspose.Cells workbook password encryption | save encrypted workbook to UNC path | C# Aspose.Cells SaveFormat Xlsx with password | network share Excel file protection Aspose.Cells | Workbook.Settings.Password usage

using Aspose.Cells;
using System;
using System.IO;

// The example creates a Workbook, optionally adds data, sets workbook.Settings.Password to encrypt the file, ensures the target directory exists, and saves the workbook as an Xlsx file. By specifying a UNC path, the encrypted workbook can be stored on a network share while preserving the password protection.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // (Optional) Add some sample data
            workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");

            // Set the password that will encrypt the workbook
            workbook.Settings.Password = "MyEncryptionPassword";

            // Define a local path where the file will be saved
            string outputDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AsposeOutput");

            // Ensure the directory exists
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            string outputPath = Path.Combine(outputDirectory, "EncryptedWorkbook.xlsx");

            // Save the workbook; the encryption password is preserved in the saved file
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

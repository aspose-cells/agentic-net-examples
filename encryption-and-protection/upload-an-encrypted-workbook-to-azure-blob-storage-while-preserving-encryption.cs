// Title: Upload an encrypted Excel workbook to Azure Blob Storage while preserving its password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# upload password‑protected .xlsx to Azure Blob storage | How to keep Excel workbook encryption when saving to Azure Blob with Aspose.Cells | Upload encrypted workbook to Azure Blob using MemoryStream in .NET | Preserve Excel file password during cloud storage upload with Aspose.Cells
// Tags: Aspose.Cells upload encrypted Excel to Azure Blob | C# preserve workbook password during cloud upload | LoadOptions password protected XLSX Aspose.Cells | OoxmlSaveOptions encrypted workbook .NET | MemoryStream Azure Blob upload Aspose.Cells | Azure Blob storage encrypted Excel handling

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads a password‑protected Excel file with Aspose.Cells LoadOptions, reassigns the password to the workbook settings to retain encryption, saves the workbook to a MemoryStream using OoxmlSaveOptions, and then streams the data to Azure Blob Storage (simulated locally), ensuring the workbook remains encrypted throughout the upload process.
class Program
{
    static async Task Main(string[] args)
    {
        // Path to the local encrypted workbook
        string localEncryptedFilePath = "encrypted.xlsx";

        // Password used to encrypt the workbook
        string workbookPassword = "Secret123";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(localEncryptedFilePath))
        {
            Console.WriteLine($"Error: File '{localEncryptedFilePath}' not found.");
            return;
        }

        try
        {
            // Load the encrypted workbook while preserving its encryption
            using (FileStream fileStream = File.OpenRead(localEncryptedFilePath))
            {
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = workbookPassword
                };
                var workbook = new Workbook(fileStream, loadOptions);

                // Ensure the workbook remains encrypted when saved
                workbook.Settings.Password = workbookPassword;

                // Save the workbook to a memory stream
                using (var memoryStream = new MemoryStream())
                {
                    var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
                    workbook.Save(memoryStream, saveOptions);
                    memoryStream.Position = 0; // Reset stream position for further use

                    // Simulate upload by saving locally
                    string simulatedUploadPath = "uploaded_encrypted.xlsx";
                    using (var fileOut = new FileStream(simulatedUploadPath, FileMode.Create, FileAccess.Write))
                    {
                        await memoryStream.CopyToAsync(fileOut);
                    }

                    Console.WriteLine($"Encrypted workbook saved locally as '{simulatedUploadPath}'.");
                }
            }

            Console.WriteLine("Operation completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

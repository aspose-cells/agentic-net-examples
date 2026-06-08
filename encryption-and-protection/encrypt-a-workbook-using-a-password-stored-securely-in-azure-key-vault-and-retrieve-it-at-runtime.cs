using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Retrieve password (replace with Azure Key Vault integration if needed)
            string password = GetPassword();

            // Create a new workbook and add data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue(
                "This workbook is encrypted with a password retrieved securely.");

            // Apply password protection
            workbook.Settings.Password = password;

            // Save the encrypted workbook
            string outputPath = "EncryptedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");

            // Verify the file exists before loading
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("Encrypted workbook not found.", outputPath);

            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook loadedWorkbook = new Workbook(outputPath, loadOptions);
            Console.WriteLine("Loaded cell value: " + loadedWorkbook.Worksheets[0].Cells["A1"].Value);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder for password retrieval logic.
    // Replace this method's body with Azure Key Vault code when the Azure SDK is available.
    private static string GetPassword()
    {
        // Example: read from an environment variable; fallback to a default password.
        string envPassword = Environment.GetEnvironmentVariable("WORKBOOK_PASSWORD");
        return string.IsNullOrEmpty(envPassword) ? "DefaultPassword123!" : envPassword;
    }
}
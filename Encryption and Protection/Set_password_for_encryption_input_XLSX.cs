using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the existing XLSX file
        string inputPath = "input.xlsx";

        // Path where the password‑protected workbook will be saved
        string outputPath = "output_protected.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Set the password that will be required to open the workbook
        workbook.Settings.Password = "MySecretPassword";

        // (Optional) Define encryption strength – strong provider with 128‑bit key
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the workbook with the applied password protection
        workbook.Save(outputPath);

        // -----------------------------------------------------------------
        // Verify that the workbook is indeed protected by loading it with the password
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "MySecretPassword";

        Workbook protectedWorkbook = new Workbook(outputPath, loadOptions);

        // Output a cell value to confirm successful loading
        Console.WriteLine("Cell A1 value after loading protected workbook: " +
                          protectedWorkbook.Worksheets[0].Cells["A1"].Value?.ToString());
    }
}
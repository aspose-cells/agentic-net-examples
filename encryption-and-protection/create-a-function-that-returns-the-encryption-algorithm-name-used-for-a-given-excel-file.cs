// Title: C# function to obtain the encryption algorithm name of an Excel file using Aspose.Cells
// Common Searches: How can I get the encryption algorithm of an encrypted Excel workbook with Aspose.Cells in C#? | Aspose.Cells C# sample to detect which encryption method is applied to an .xlsx file | Retrieve Excel file protection algorithm using LoadOptions in Aspose.Cells .NET | Determine if an Excel file is encrypted and identify the algorithm via Aspose.Cells | C# code to read EncryptionAlgorithm property of a password‑protected workbook with Aspose.Cells
// Tags: Aspose.Cells LoadOptions EncryptionAlgorithm C# | detect Excel workbook encryption algorithm .NET | read encrypted .xlsx protection type Aspose.Cells | C# function get Excel file encryption method | Aspose.Cells encrypted workbook detection

using System;
using System.IO;
using Aspose.Cells;

// The example defines a static C# method GetEncryptionAlgorithm that validates the file path, loads the workbook with Aspose.Cells LoadOptions (without a password), and uses reflection to read the LoadOptions.EncryptionAlgorithm property. It returns the algorithm name, "None" when the workbook is not encrypted, or "Unknown" if the encryption cannot be identified.
public class ExcelEncryptionHelper
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>Name of the encryption algorithm.</returns>
    public static string GetEncryptionAlgorithm(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"The file '{filePath}' does not exist.");

        try
        {
            // LoadOptions without a password allows Aspose.Cells to detect encryption.
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook; if encrypted, Aspose.Cells still loads it and fills EncryptionAlgorithm.
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Retrieve the EncryptionAlgorithm property via reflection for compatibility.
            var algorithmProperty = typeof(LoadOptions).GetProperty("EncryptionAlgorithm");
            if (algorithmProperty != null)
            {
                var algorithmValue = algorithmProperty.GetValue(loadOptions);
                return algorithmValue?.ToString() ?? "None";
            }

            // Property not available – assume no encryption.
            return "None";
        }
        catch (CellsException)
        {
            // The file is encrypted but the algorithm could not be determined.
            return "Unknown";
        }
        catch (Exception)
        {
            // Propagate other exceptions.
            throw;
        }
    }

    // Simple entry point for demonstration purposes.
    public static void Main(string[] args)
    {
        try
        {
            string path = args.Length > 0 ? args[0] : "sample.xlsx";
            string algorithm = GetEncryptionAlgorithm(path);
            Console.WriteLine($"Encryption algorithm for '{path}': {algorithm}");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine(fnfEx.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

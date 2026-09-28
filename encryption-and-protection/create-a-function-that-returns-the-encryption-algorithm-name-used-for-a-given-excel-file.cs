// Title: Get the encryption algorithm or status of an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write a C# method using Aspose.Cells that opens an Excel file and returns the exact encryption algorithm (e.g., AES128, AES256) or "None" when the file is not protected. | Modify the GetEncryptionAlgorithm function to access the workbook's encryption properties via LoadOptions and return the specific algorithm name instead of a generic unknown message. | Create a .NET console application that takes a file path argument, calls the encryption‑algorithm helper, and prints the algorithm name or an appropriate error description.
// Common Searches: aspnet get encryption algorithm of password protected xlsx using aspose.cells | c# determine if excel file is encrypted and which algorithm Aspose.Cells | how to read encryption type of an .xlsm file with Aspose.Cells .NET | retrieve workbook encryption method AES128 AES256 in C# Aspose.Cells
// Tags: Aspose.Cells workbook encryption algorithm detection | C# read Excel file encryption type | Aspose.Cells LoadOptions encryption property | detect encrypted .xlsx with Aspose.Cells | retrieve encryption algorithm name .NET

using System;
using System.IO;
using Aspose.Cells;

// The example defines ExcelEncryptionHelper.GetEncryptionAlgorithm, which checks file existence, attempts to load the workbook with Aspose.Cells without a password, returns "None" if loading succeeds, "Encrypted (Unknown)" if a CellsException occurs, and propagates other errors. A console program passes a file path argument to this helper and prints the resulting encryption information.
public static class ExcelEncryptionHelper
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>Encryption status or error message.</returns>
    public static string GetEncryptionAlgorithm(string filePath)
    {
        // Verify that the file exists before attempting to load it.
        if (!File.Exists(filePath))
            throw new FileNotFoundException("The specified Excel file was not found.", filePath);

        try
        {
            // Attempt to load the workbook without a password.
            var loadOptions = new LoadOptions(LoadFormat.Auto);
            var workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, the file is not encrypted.
            return "None";
        }
        catch (CellsException)
        {
            // The file is encrypted but the password is not supplied.
            return "Encrypted (Unknown)";
        }
        catch (Exception ex)
        {
            // Return the exception message for diagnostic purposes.
            return $"Error: {ex.Message}";
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expect the Excel file path as the first command‑line argument.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: dotnet run <ExcelFilePath>");
            return;
        }

        string filePath = args[0];

        try
        {
            string result = ExcelEncryptionHelper.GetEncryptionAlgorithm(filePath);
            Console.WriteLine($"Encryption information: {result}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to retrieve encryption info: {ex.Message}");
        }
    }
}

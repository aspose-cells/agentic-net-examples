// Title: Identify the encryption status of an Excel workbook using Aspose.Cells in C#
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells in C# | Aspose.Cells C# detect encrypted workbook without password | retrieve encryption algorithm name from protected .xlsx with Aspose.Cells | C# code to determine encryption status of an Excel workbook using Aspose.Cells | Aspose.Cells get encryption type of Excel file programmatically
// Tags: Aspose.Cells workbook encryption check C# | C# identify Excel protection status Aspose.Cells | Aspose.Cells retrieve encryption algorithm indicator | C# load Excel file without password Aspose.Cells | Aspose.Cells handling encrypted Excel files

using System;
using System.IO;
using Aspose.Cells;

// The GetEncryptionAlgorithm method validates the file path, attempts to load the workbook with Aspose.Cells without a password, returns "None" when loading succeeds (indicating no encryption), and returns "Encrypted" when a CellsException is caught, signaling that the Excel file is protected.
public class ExcelEncryptionHelper
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>Name of the encryption algorithm or "None".</returns>
    public static string GetEncryptionAlgorithm(string filePath)
    {
        // Ensure the file exists before attempting to load it.
        if (!File.Exists(filePath))
            throw new FileNotFoundException("The specified Excel file was not found.", filePath);

        try
        {
            // Attempt to load the workbook without a password.
            // If the file is not encrypted this will succeed.
            Workbook workbook = new Workbook(filePath);
            // Aspose.Cells does not expose the algorithm directly; indicate no encryption.
            return "None";
        }
        catch (CellsException)
        {
            // An exception here typically means the workbook is encrypted.
            return "Encrypted";
        }
        catch (Exception)
        {
            // Propagate unexpected exceptions.
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the full path to the Excel file as an argument.");
                return;
            }

            string filePath = args[0];
            string algorithm = ExcelEncryptionHelper.GetEncryptionAlgorithm(filePath);
            Console.WriteLine($"Encryption algorithm: {algorithm}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

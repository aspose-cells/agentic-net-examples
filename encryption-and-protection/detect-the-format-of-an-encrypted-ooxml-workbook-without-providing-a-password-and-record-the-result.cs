// Title: Detect Encryption of an OOXML (.xlsx) Workbook with Aspose.Cells (C#) Without a Password
// Description: A C# example that uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the file format of an .xlsx workbook and determine whether it is encrypted, all without providing a password. The program writes the detection outcome to the console and saves it to a text file, with robust handling for missing files and unexpected errors.
// Keywords: Aspose.Cells | C# | detect encrypted workbook | FileFormatUtil | OOXML | xlsx encryption detection | without password | file format detection | Excel password check | log encryption status
// Common Searches: How to check if an .xlsx file is password protected using Aspose.Cells C# | Detect encrypted Excel workbook without opening it | FileFormatUtil DetectFileFormat encryption status example | Save Excel encryption detection result to a file | Aspose.Cells detect encrypted workbook in .NET
// Developer Intent: Determine whether an OOXML Excel workbook is encrypted without supplying a password and record the result for downstream processing.
// Use Cases: Validate incoming spreadsheet uploads and reject encrypted files before further processing. | Create audit logs that capture the encryption status of batch‑processed Excel files. | Route encrypted and unencrypted workbooks to separate workflows in an automated pipeline.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect if an .xlsx file is password protected without opening the workbook, and export the result to JSON. | Show how to handle exceptions from FileFormatUtil.DetectFileFormat when the file is corrupted or in an unsupported format. | Provide a sample ASP.NET Core controller that validates uploaded Excel files for encryption using Aspose.Cells and returns a clear error message if encrypted.

using System;
using System.IO;
using Aspose.Cells;

// A C# example that uses Aspose.Cells FileFormatUtil.DetectFileFormat to identify the file format of an .xlsx workbook and determine whether it is encrypted, all without providing a password. The program writes the detection outcome to the console and saves it to a text file, with robust handling for missing files and unexpected errors.
public class DetectEncryptedWorkbook
{
    public static void Run()
    {
        // Path to the OOXML workbook (encrypted or not)
        string filePath = "encrypted.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Detect the file format without providing a password
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Output detection results to the console
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Detected Format: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

            // Record the results to a text file
            string resultPath = "DetectionResult.txt";
            using (StreamWriter writer = new StreamWriter(resultPath, false))
            {
                writer.WriteLine($"File: {filePath}");
                writer.WriteLine($"Detected Format: {formatInfo.FileFormatType}");
                writer.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
            }

            Console.WriteLine($"Detection results saved to {resultPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., format detection issues)
            Console.WriteLine($"An error occurred during detection: {ex.Message}");
        }
    }
}

// Entry point required for the application
public class Program
{
    public static void Main(string[] args)
    {
        DetectEncryptedWorkbook.Run();
    }
}

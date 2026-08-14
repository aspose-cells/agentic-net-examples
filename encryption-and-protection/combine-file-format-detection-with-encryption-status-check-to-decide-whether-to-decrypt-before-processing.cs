// Title: Detect Excel format & encryption, then conditionally decrypt with Aspose.Cells (C#)
// Description: C# sample that uses Aspose.Cells FileFormatUtil to identify a spreadsheet's format and encryption flag, loads the workbook with a password only when needed, auto‑fits the first worksheet's columns, and saves the result as a new file.
// Keywords: Aspose.Cells file format detection | Excel encryption check C# | load encrypted workbook Aspose | conditional decryption Aspose.Cells | .NET auto fit columns | process mixed‑format spreadsheets | sample code GitHub
// Common Searches: detect password protected Excel with Aspose.Cells | load encrypted workbook using LoadOptions password | auto detect Excel format before opening | C# example for conditional decryption of spreadsheets | Aspose.Cells sample for mixed file types
// Developer Intent: Identify a spreadsheet's type and encryption state, then open it with the appropriate credentials before applying any processing.
// Use Cases: Open user‑uploaded Excel files of unknown type, automatically handle password‑protected workbooks, and apply formatting changes. | Batch‑process a folder containing XLS, XLSX, CSV, and encrypted files without manual format checks. | Integrate format detection and conditional decryption into a web API that returns a cleaned version of the workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect an Excel file's format and encryption status, then loads it with a password only if encrypted and saves an auto‑fitted copy. | Create a reusable method accepting a file path and optional password, performing format detection, conditional decryption, and column auto‑fit on the first worksheet. | Explain best practices for handling missing files, wrong passwords, and other exceptions when loading encrypted workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// C# sample that uses Aspose.Cells FileFormatUtil to identify a spreadsheet's format and encryption flag, loads the workbook with a password only when needed, auto‑fits the first worksheet's columns, and saves the result as a new file.
public class DetectAndDecryptDemo
{
    // Demonstrates detection of file format and encryption status,
    // then loads the workbook with or without a password accordingly.
    public static void Run(string filePath, string password)
    {
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found – \"{filePath}\"");
            return;
        }

        try
        {
            // Detect file format and whether the file is encrypted
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is encrypted: {formatInfo.IsEncrypted}");

            Workbook workbook;

            if (formatInfo.IsEncrypted)
            {
                // File is encrypted – load it using the supplied password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = password
                };
                workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded with password.");
            }
            else
            {
                // File is not encrypted – load normally
                workbook = new Workbook(filePath);
                Console.WriteLine("Workbook loaded without password.");
            }

            // Example processing: auto‑fit columns of the first worksheet
            if (workbook.Worksheets.Count > 0)
            {
                workbook.Worksheets[0].AutoFitColumns();
            }

            // Save the processed workbook to a new file
            string outputPath = Path.Combine(
                Path.GetDirectoryName(filePath) ?? string.Empty,
                "processed_" + Path.GetFileName(filePath));
            workbook.Save(outputPath);
            Console.WriteLine($"Processed workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    // Entry point required for compilation
    public static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: <exe> <filePath> [password]");
            return;
        }

        string filePath = args[0];
        string password = args.Length >= 2 ? args[1] : string.Empty;

        DetectAndDecryptDemo.Run(filePath, password);
    }
}

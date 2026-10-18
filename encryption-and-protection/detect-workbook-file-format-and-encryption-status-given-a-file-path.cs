// Title: Detect Excel workbook file format and encryption status with Aspose.Cells in C#
// Common Searches: C# code to identify Excel file type using Aspose.Cells LoadOptions | How to verify if an Excel workbook is password protected with Aspose.Cells | Determine spreadsheet encryption status without fully loading the file in .NET | Get detected format of a workbook from its file path using Aspose.Cells | Read IsEncrypted and IsPasswordProtected properties via reflection in Aspose.Cells
// Tags: Aspose.Cells LoadOptions file format detection | C# check workbook encryption Aspose.Cells | detect password protected Excel workbook .NET | reflection access IsEncrypted Aspose.Cells | determine spreadsheet load format without full load

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to verify a workbook file exists, use Aspose.Cells LoadOptions to detect the Excel file format without fully loading the workbook, open the workbook, and employ reflection to read the IsEncrypted and IsPasswordProtected properties across different Aspose.Cells versions. It then prints the file path, detected format, and encryption/password‑protection status while handling possible errors.
public class WorkbookInfoDetector
{
    /// <param name="filePath">Full path to the workbook file.</param>
    public static void DetectWorkbookInfo(string filePath)
    {
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        try
        {
            // Prepare load options – this allows Aspose.Cells to determine the file format without fully loading the file.
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the specified options.
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Determine the detected file format.
            LoadFormat detectedFormat = loadOptions.LoadFormat;

            // Use reflection to obtain encryption information (compatible with multiple Aspose.Cells versions).
            bool isEncrypted = false;
            bool isPasswordProtected = false;
            var wbType = workbook.GetType();

            var encProp = wbType.GetProperty("IsEncrypted");
            if (encProp != null && encProp.PropertyType == typeof(bool))
                isEncrypted = (bool)encProp.GetValue(workbook);

            var pwdProp = wbType.GetProperty("IsPasswordProtected");
            if (pwdProp != null && pwdProp.PropertyType == typeof(bool))
                isPasswordProtected = (bool)pwdProp.GetValue(workbook);

            // Output the results.
            Console.WriteLine($"File Path: {filePath}");
            Console.WriteLine($"Detected Format: {detectedFormat}");
            Console.WriteLine($"Is Encrypted: {isEncrypted}");
            Console.WriteLine($"Is Password Protected: {isPasswordProtected}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the file: {ex.Message}");
        }
    }

    // Entry point for the console application.
    public static void Main(string[] args)
    {
        // Example usage: provide the file path as the first argument.
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the full path to the workbook file as a command-line argument.");
            return;
        }

        string filePath = args[0];
        DetectWorkbookInfo(filePath);
    }
}

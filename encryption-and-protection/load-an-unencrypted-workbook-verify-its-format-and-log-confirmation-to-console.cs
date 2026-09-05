// Title: Load an unencrypted Excel workbook and output its format type and version with Aspose.Cells for .NET
// AI Prompts: Open a specified .xlsx file using Aspose.Cells Workbook and print a success message to the console. | Invoke FileFormatInfo.GetFileFormatInfo via reflection to obtain the workbook's FileFormatType and Version, then log the values. | Verify the file exists and handle the case where FileFormatInfo is unavailable, displaying appropriate error messages.
// Common Searches: aspocells get file format type of an Excel workbook programmatically c# | how to retrieve Excel file version using Aspose.Cells without encryption | using reflection to call FileFormatInfo in older Aspose.Cells releases | load unencrypted workbook and display format information c# aspocells | detect if Excel file is encrypted before loading with Aspose.Cells
// Tags: open unencrypted Excel file Aspose.Cells | retrieve Excel file format type via FileFormatInfo | reflection based FileFormatInfo access Aspose.Cells | determine workbook version programmatically .NET | fallback handling when FileFormatInfo unavailable

using System;
using System.IO;
using Aspose.Cells;
using System.Reflection;

// The example checks that a given .xlsx file exists, loads it with Aspose.Cells Workbook, attempts to fetch the file's format type and version using FileFormatInfo through reflection, and writes success, format details, or error information to the console.
class Program
{
    static void Main()
    {
        // Path to the workbook file (replace with your actual file path)
        string workbookPath = "sample.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook (unencrypted)
            Workbook workbook = new Workbook(workbookPath);
            Console.WriteLine("Workbook loaded successfully.");

            // Try to obtain format information using FileFormatInfo if the API is available
            var formatInfo = GetFileFormatInfo(workbookPath);
            if (formatInfo != null)
            {
                Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");
                Console.WriteLine($"Version: {formatInfo.Version}");
            }
            else
            {
                Console.WriteLine("FileFormatInfo API not available in this Aspose.Cells version.");
            }
        }
        catch (Exception ex)
        {
            // Log any errors that occur during loading or verification
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }

    // Uses reflection to call FileFormatInfo.GetFileFormatInfo(string) if it exists.
    private static dynamic GetFileFormatInfo(string path)
    {
        try
        {
            Type fileFormatInfoType = typeof(FileFormatInfo);
            MethodInfo method = fileFormatInfoType.GetMethod("GetFileFormatInfo", new[] { typeof(string) });
            if (method != null)
            {
                return method.Invoke(null, new object[] { path });
            }
        }
        catch
        {
            // Suppress any reflection errors; return null to indicate unavailability.
        }
        return null;
    }
}

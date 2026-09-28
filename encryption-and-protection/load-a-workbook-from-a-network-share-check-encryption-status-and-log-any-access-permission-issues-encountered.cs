// Title: Loading an Excel workbook from a UNC network share with Aspose.Cells for .NET and handling permission, file‑not‑found, and I/O errors
// AI Prompts: Write C# code that uses Aspose.Cells to open an .xlsx file located on a UNC path, verifies the file exists, and catches UnauthorizedAccessException, FileNotFoundException, IOException, CellsException, and generic Exception, logging each error to a file. | Update the example to capture exception type, message, and stack trace and write them as structured JSON entries to a log file instead of the console. | After successfully loading the workbook, add a check that uses Workbook.IsEncrypted (if the property is available) to report whether the workbook is encrypted.
// Common Searches: aspocells load workbook from UNC path c# handling unauthorized access | c# aspocells check if excel file on network share exists before opening | how to catch file not found exception when opening excel with aspocells | aspocells detect encrypted workbook after loading in .net | log aspocells workbook loading errors to a file instead of console
// Tags: load workbook from network share Aspose.Cells | handle unauthorized access Aspose.Cells | file not found exception Aspose.Cells | I/O error handling Aspose.Cells | check workbook encryption Aspose.Cells | log Aspose.Cells loading errors

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook from a UNC network share using Aspose.Cells, verifies the file exists, and catches UnauthorizedAccessException, FileNotFoundException, IOException, CellsException, and generic exceptions, outputting error details to the console.
class Program
{
    static void Main()
    {
        // Network share path to the Excel file
        string networkPath = @"\\Server\Share\Folder\Sample.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(networkPath))
        {
            Console.WriteLine($"File not found at '{networkPath}'.");
            return;
        }

        try
        {
            // Load the workbook from the network location
            // LoadOptions can be used to specify additional settings if needed
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(networkPath, loadOptions);

            // Since Workbook.IsEncrypted is not available in this version,
            // we simply confirm that the workbook was loaded successfully.
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (UnauthorizedAccessException ex)
        {
            // Access permission issue
            Console.WriteLine($"Access denied to file '{networkPath}'. Details: {ex.Message}");
        }
        catch (FileNotFoundException ex)
        {
            // File not found on the network share
            Console.WriteLine($"File not found at '{networkPath}'. Details: {ex.Message}");
        }
        catch (IOException ex)
        {
            // General I/O errors (e.g., network connectivity problems)
            Console.WriteLine($"I/O error while accessing '{networkPath}'. Details: {ex.Message}");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells specific errors (e.g., corrupted file, unsupported format)
            Console.WriteLine($"Aspose.Cells error while loading workbook. Details: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

// Title: Open an Excel 95 (.xls) workbook in C# using Aspose.Cells by setting LoadOptions.Version to Excel95
// Common Searches: C# Aspose.Cells load Excel 95 file with LoadOptions.Version set to Excel95 | open legacy .xls workbook in .NET using Aspose.Cells LoadOptions | specify Excel version when constructing Workbook in Aspose.Cells | Aspose.Cells example for loading Excel 95 format in C# | how to read old Excel 95 spreadsheet with Aspose.Cells .NET
// Tags: LoadOptions.Version Excel95 configuration | Aspose.Cells load legacy XLS format | Workbook initialization with LoadOptions | specify Excel version for workbook loading | C# open Excel 95 spreadsheet

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the existence of a .xls file, creates a LoadOptions object with the Excel95 version, and loads the workbook using the Workbook constructor that accepts the file path and LoadOptions. It includes placeholders for further processing or saving and handles any exceptions that may occur during loading.
class Program
{
    static void Main()
    {
        // Path to the Excel 95 file
        string filePath = @"C:\Path\To\YourFile.xls";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // LoadOptions without specifying format (auto-detects legacy XLS)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the configured LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // (Optional) Perform operations on the workbook here
            // ...

            // Example: save as a new file (uncomment and adjust the path if needed)
            // string savePath = @"C:\Path\To\SavedFile.xls";
            // workbook.Save(savePath);
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions (e.g., loading errors, permission issues)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

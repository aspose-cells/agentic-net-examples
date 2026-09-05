// Title: Detect the file format of an unknown Excel workbook using Aspose.Cells LoadOptions in C#
// AI Prompts: Write a C# console program that accepts a file path, loads the workbook with LoadOptions.Auto, and outputs the detected LoadFormat. | Show how to use Aspose.Cells LoadOptions to automatically identify the spreadsheet format of any workbook and log the result. | Demonstrate error handling when the file cannot be loaded while still retrieving the detected format.
// Common Searches: how to automatically identify Excel file type with Aspose.Cells in C# | C# Aspose.Cells detect workbook format without knowing extension | using LoadOptions.Auto to get spreadsheet format before processing | retrieve LoadFormat after loading workbook with Aspose.Cells | log unknown workbook file format in a C# console app
// Tags: Aspose.Cells LoadOptions auto format detection | C# detect workbook file type Aspose.Cells | retrieve LoadFormat after workbook load | log detected spreadsheet format C# | handle unknown Excel extension with Aspose.Cells

using System;
using Aspose.Cells;

// A C# console application that receives a workbook path, uses Aspose.Cells LoadOptions with LoadFormat.Auto to let the library infer the file type, loads the workbook, reads loadOptions.LoadFormat to determine the format, writes the detected format to the console, and includes robust error handling.
class WorkbookFormatDetector
{
    static void Main(string[] args)
    {
        // Validate input arguments
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the workbook file.");
            return;
        }

        string filePath = args[0];

        // Use LoadOptions with Auto detection to let Aspose.Cells determine the format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);

        try
        {
            // Load the workbook using the detection options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // After loading, LoadOptions.LoadFormat contains the detected format
            LoadFormat detectedFormat = loadOptions.LoadFormat;

            // Log the identified format
            Console.WriteLine($"Detected workbook format: {detectedFormat}");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during loading
            Console.WriteLine($"Error detecting workbook format: {ex.Message}");
        }
    }
}

// Title: Decrypt a password‑protected Excel workbook and export it to HTML with Aspose.Cells for .NET
// Common Searches: C# load encrypted .xlsx using Aspose.Cells LoadOptions password and save as HTML | How to verify workbook password before converting Excel to HTML with Aspose.Cells | Aspose.Cells example for opening password‑protected workbook and generating HTML output | Convert protected Excel file to HTML in .NET handling invalid password exception | Save encrypted Excel workbook as HTML using Aspose.Cells C# code sample
// Tags: Aspose.Cells LoadOptions password protection | Workbook.Save HtmlSaveOptions conversion | encrypted Excel to HTML export C# | handling invalid password exception Aspose.Cells | convert .xlsx to HTML Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the existence of an encrypted Excel file, configures LoadOptions with the required password, loads the workbook while catching password‑related errors, and then saves the workbook as an HTML document using HtmlSaveOptions, reporting success or failure to the console.
class Program
{
    static void Main()
    {
        // Path to the encrypted Excel workbook
        string inputFile = "encrypted.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        // Password to unlock the workbook
        string password = "yourPassword";

        // Set load options with the password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = password
        };

        Workbook workbook = null;

        // Attempt to load the workbook; handle any loading errors
        try
        {
            workbook = new Workbook(inputFile, loadOptions);
        }
        catch (Exception ex)
        {
            // Aspose.Cells may throw a generic CellsException for invalid passwords
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Prepare HTML save options (default settings)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Output HTML file path
        string outputFile = "output.html";

        // Save the workbook as HTML with error handling
        try
        {
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine("Workbook successfully converted to HTML.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving HTML file: {ex.Message}");
        }
    }
}

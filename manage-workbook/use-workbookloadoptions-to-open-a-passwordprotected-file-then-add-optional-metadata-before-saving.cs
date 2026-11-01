// Title: Load a password‑protected XLSX workbook with LoadOptions, add custom document properties, and save using OoxmlSaveOptions in C# (Aspose.Cells)
// Common Searches: C# load a password‑protected Excel file with Aspose.Cells LoadOptions | add custom document properties to an XLSX workbook using Aspose.Cells | save a workbook with metadata using OoxmlSaveOptions in .NET | example of opening a protected workbook and setting new properties in C# | Aspose.Cells how to set and read custom document properties after loading a file
// Tags: LoadOptions protected XLSX Aspose.Cells | custom document properties API Aspose.Cells | OoxmlSaveOptions metadata saving | Aspose.Cells workbook metadata handling | C# load workbook with password using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

// The sample verifies the input file, loads a password‑protected XLSX workbook using LoadOptions, adds three custom document properties (Author, Reviewed, ReviewDate), configures OoxmlSaveOptions (with optional new password), and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "protected.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Set load options with the password for the protected file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "myPassword"
            };

            // Load the password‑protected workbook
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Add optional metadata (custom document properties)
            workbook.CustomDocumentProperties.Add("Author", "John Doe");
            workbook.CustomDocumentProperties.Add("Reviewed", true);
            workbook.CustomDocumentProperties.Add("ReviewDate", DateTime.Now);

            // Prepare save options (use OoxmlSaveOptions for XLSX format)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            // saveOptions.Password = "newPassword"; // Uncomment to protect the saved file

            // Save the workbook with the added metadata
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

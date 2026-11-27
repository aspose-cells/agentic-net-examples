// Title: Create a new XLSX workbook in C# and protect it with a strong opening password using Aspose.Cells
// Common Searches: how to add an opening password to a newly created Excel file with Aspose.Cells in C# | set strong encryption for an XLSX workbook when saving with Aspose.Cells .NET | save a password‑protected workbook using OoxmlSaveOptions in C# | ensure the output directory exists before saving a protected Excel file with Aspose.Cells | apply workbook.Settings.Password for strong password protection in Aspose.Cells
// Tags: Aspose.Cells workbook opening password | C# encrypt XLSX with OoxmlSaveOptions | strong password protection Aspose.Cells | create directory before workbook save | password‑protected Excel file .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a new Workbook, assigns a strong opening password via workbook.Settings.Password, configures OoxmlSaveOptions for XLSX format, ensures the target directory exists, and saves the file as a password‑protected Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set a password required to open the workbook
            workbook.Settings.Password = "StrongP@ssw0rd!";

            // Configure save options (encryption via password is handled by workbook settings)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Define output file path
            string outputPath = "ProtectedWorkbook.xlsx";

            // Ensure the directory exists (handle possible null directory)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook with the password applied
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

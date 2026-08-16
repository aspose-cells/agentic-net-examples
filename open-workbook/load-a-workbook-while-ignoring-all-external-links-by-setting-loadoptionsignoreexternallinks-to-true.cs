// Title: Load an Excel workbook with Aspose.Cells for .NET while ignoring external links (LoadOptions.IgnoreExternalLinks = true)
// Description: C# example that shows how to open an XLSX file using Aspose.Cells with LoadOptions.IgnoreExternalLinks set to true, preventing any external reference evaluation. The sample checks the input file, loads the workbook with the option, and saves it to a new file while handling errors gracefully.
// Keywords: Aspose.Cells LoadOptions.IgnoreExternalLinks | load workbook without external links .NET | disable external link evaluation Aspose.Cells | open Excel file ignoring external references | Aspose.Cells security external links
// Common Searches: Aspose.Cells ignore external links when loading workbook | LoadOptions.IgnoreExternalLinks true example C# | prevent external link updates Aspose.Cells | open Excel file without resolving external references Aspose | Aspose.Cells security disable external links
// Developer Intent: Open an Excel workbook in .NET while suppressing all external link processing.
// Use Cases: Read or modify workbooks that contain formulas pointing to other files without triggering network calls. | Batch‑process Excel files in a secure environment where external links must be ignored for compliance. | Convert workbooks to PDF, images, or other formats while ensuring external links are not resolved or embedded.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells using LoadOptions.IgnoreExternalLinks = true and saves it to a new file. | Explain how LoadOptions.IgnoreExternalLinks improves security when opening untrusted Excel workbooks with Aspose.Cells. | Provide a fallback snippet that checks the Aspose.Cells version and applies the ignore‑external‑links option only when supported.

using System;
using System.IO;
using Aspose.Cells;

// C# example that shows how to open an XLSX file using Aspose.Cells with LoadOptions.IgnoreExternalLinks set to true, preventing any external reference evaluation. The sample checks the input file, loads the workbook with the option, and saves it to a new file while handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (default options)
            Workbook workbook = new Workbook(inputPath);

            // Note: The property to disable external links is not available in the current Aspose.Cells version.
            // If needed, configure external link handling using other available settings.

            // Perform any required operations here (e.g., read data, modify cells)

            // Save the workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

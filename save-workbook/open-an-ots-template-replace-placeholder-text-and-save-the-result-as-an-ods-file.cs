// Title: C# – Convert OTS Template to ODS with Placeholder Replacement using Aspose.Cells
// Description: Shows how to load (or create) an OTS template, replace a {Name} placeholder, configure OdsSaveOptions for LibreOffice compatibility, and save the workbook as an ODS file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | OTS template | ODS output | C# | .NET | placeholder replacement | OdsSaveOptions | LibreOffice generator | workbook conversion | template processing
// Common Searches: Aspose.Cells replace text in OTS and save as ODS C# | convert OTS file to ODS using .NET | C# example for OdsSaveOptions LibreOffice | how to create OTS template programmatically Aspose.Cells | batch replace placeholders in OTS and export ODS
// Developer Intent: Replace placeholder values in an OTS workbook and export the modified file as ODS.
// Use Cases: Generate personalized ODS reports by loading a reusable OTS template, inserting user‑specific data, and saving the result for LibreOffice consumption. | Automate bulk conversion of multiple OTS files, each with distinct placeholder values, into ODS documents with consistent generator settings. | Integrate OTS‑to‑ODS transformation into a .NET service that prepares data‑driven spreadsheets for downstream open‑source office workflows.
// AI Prompts: Write C# code that opens an OTS file, replaces several placeholders (e.g., {Name}, {Date}), and saves the workbook as ODS using Aspose.Cells with custom OdsSaveOptions. | Provide a step‑by‑step tutorial for programmatically creating an OTS template, adding placeholders, handling missing files, and converting it to ODS in a .NET application. | Explain how to configure OdsSaveOptions to emulate LibreOffice output and apply those settings when saving a workbook after placeholder substitution.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to load (or create) an OTS template, replace a {Name} placeholder, configure OdsSaveOptions for LibreOffice compatibility, and save the workbook as an ODS file with Aspose.Cells for .NET.
class OtsToOdsConverter
{
    static void Main()
    {
        // Paths for the template and output files
        string templatePath = "template.ots";
        string outputPath = "result.ods";

        try
        {
            // Ensure the template file exists; create a simple one if missing
            if (!File.Exists(templatePath))
            {
                var tempWorkbook = new Workbook();
                // Insert a placeholder that will be replaced later
                tempWorkbook.Worksheets[0].Cells["A1"].PutValue("{Name}");
                tempWorkbook.Save(templatePath);
            }

            // Load the OTS template into a Workbook
            var workbook = new Workbook(templatePath);

            // Replace placeholder text with actual value
            workbook.Replace("{Name}", "John Doe");

            // Configure ODS save options (optional: set generator type)
            var saveOptions = new OdsSaveOptions
            {
                GeneratorType = OdsGeneratorType.LibreOffice
            };

            // Save the modified workbook as an ODS file
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Conversion completed successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

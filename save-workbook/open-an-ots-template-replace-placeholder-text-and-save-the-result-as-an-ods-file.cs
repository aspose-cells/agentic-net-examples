using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace OtsToOdsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the OTS template file
            string templatePath = "Template.ots";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file not found at '{Path.GetFullPath(templatePath)}'.");
                return;
            }

            try
            {
                // Load the OTS template using OdsLoadOptions
                OdsLoadOptions loadOptions = new OdsLoadOptions();
                Workbook workbook = new Workbook(templatePath, loadOptions);

                // Replace placeholder text in the workbook
                // Example: replace all occurrences of "{{Name}}" with "John Doe"
                workbook.Replace("{{Name}}", "John Doe");

                // Prepare ODS save options
                OdsSaveOptions saveOptions = new OdsSaveOptions
                {
                    GeneratorType = OdsGeneratorType.LibreOffice
                };

                // Save the modified workbook as an ODS file
                string outputPath = "Result.ods";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Template processed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
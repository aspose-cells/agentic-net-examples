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
            // Path to the OTS template file (relative to the executable directory)
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template.ots");

            Workbook workbook;

            if (File.Exists(templatePath))
            {
                // Load the existing OTS template
                workbook = new Workbook(templatePath);
            }
            else
            {
                // Create a new workbook and add a placeholder for demonstration
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Hello ${Name}");
            }

            // Replace placeholder text with actual value
            workbook.Replace("${Name}", "John Doe");

            // Configure ODS save options (optional: set generator type)
            OdsSaveOptions saveOptions = new OdsSaveOptions
            {
                GeneratorType = OdsGeneratorType.LibreOffice
            };

            // Save the modified workbook as an ODS file
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Result.ods");
            workbook.Save(outputPath, saveOptions);
        }
    }
}
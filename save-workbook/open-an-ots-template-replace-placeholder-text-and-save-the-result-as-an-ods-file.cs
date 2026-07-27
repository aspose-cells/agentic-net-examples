using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.ots";
            const string resultPath = "result.ods";

            // Ensure the template file exists before loading
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {Path.GetFullPath(templatePath)}");
                return;
            }

            // Load the OTS template workbook
            Workbook workbook = new Workbook(templatePath);

            // Replace placeholder texts
            workbook.Replace("{Name}", "John Doe");
            workbook.Replace("{Date}", DateTime.Now.ToString("yyyy-MM-dd"));

            // Configure ODS save options (optional)
            OdsSaveOptions saveOptions = new OdsSaveOptions
            {
                GeneratorType = OdsGeneratorType.LibreOffice
            };

            // Save the modified workbook as an ODS file
            workbook.Save(resultPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;

class ConvertToPdfFrench
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some French text to demonstrate multilingual support
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Bonjour le monde");

            // Set the workbook UI language to French (France)
            workbook.Settings.LanguageCode = CountryCode.France;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The DefaultEditLanguage property is optional.
            // It can be set if the specific enum value exists in the used Aspose.Cells version.
            // pdfOptions.DefaultEditLanguage = DefaultEditLanguage.French;

            // Define output file path
            string outputPath = "output_french.pdf";

            // Save the workbook as a PDF using the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
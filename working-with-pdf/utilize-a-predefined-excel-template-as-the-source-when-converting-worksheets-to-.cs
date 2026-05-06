using System;
using Aspose.Cells.Utility;

class ConvertTemplateToPdf
{
    static void Main()
    {
        // Path to the predefined Excel template file
        string templatePath = "Template.xlsx";

        // Desired output PDF file path
        string pdfPath = "Result.pdf";

        // Convert the Excel template to PDF using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(templatePath, pdfPath);

        Console.WriteLine("Conversion completed successfully.");
    }
}
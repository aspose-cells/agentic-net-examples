using System;
using System.IO;
using Aspose.Cells;

class CheckVbaSignatureMht
{
    static void Main()
    {
        // Load a macro-enabled workbook
        Workbook workbook = new Workbook("sample.xlsm");

        // Directly check VBA project signature using provided properties
        bool isSigned = false;
        bool isValidSigned = false;
        if (workbook.VbaProject != null)
        {
            isSigned = workbook.VbaProject.IsSigned;
            isValidSigned = workbook.VbaProject.IsValidSigned;
        }
        Console.WriteLine($"Direct check - IsSigned: {isSigned}, IsValidSigned: {isValidSigned}");

        // Save the workbook to HTML format in memory (MHTML not available in this version)
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, SaveFormat.Html);
            htmlStream.Position = 0;

            // Read the HTML content as a string
            string htmlContent = new StreamReader(htmlStream).ReadToEnd();

            // Inspect the HTML text for markers that indicate a signed VBA project
            bool htmlIndicatesSigned = htmlContent.Contains("VbaProject") && htmlContent.Contains("IsSigned=\"true\"");
            Console.WriteLine($"HTML inspection indicates signed VBA project: {htmlIndicatesSigned}");
        }
    }
}
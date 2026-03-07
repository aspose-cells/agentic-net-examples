using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string templatePath = "TemplateFile.xltx";
        string pdfPath = "ConvertedOutput.pdf";

        // Ensure the template file exists; if not, create a simple one.
        if (!File.Exists(templatePath))
        {
            Workbook tempWb = new Workbook();
            tempWb.Worksheets[0].Name = "Sheet1";
            tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            tempWb.Save(templatePath, SaveFormat.Xltx);
        }

        Workbook workbook = new Workbook(templatePath);
        workbook.Save(pdfPath, SaveFormat.Pdf);

        Console.WriteLine("Excel template has been successfully converted to PDF.");
    }
}
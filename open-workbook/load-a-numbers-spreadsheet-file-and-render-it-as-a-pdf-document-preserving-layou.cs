using System;
using Aspose.Cells;
using Aspose.Cells.Numbers;

class NumbersToPdf
{
    static void Main()
    {
        string numbersFile = "input.numbers";
        string pdfFile = "output.pdf";

        if (!System.IO.File.Exists(numbersFile))
        {
            Console.WriteLine($"Input file not found: {numbersFile}");
            return;
        }

        NumbersLoadOptions loadOptions = new NumbersLoadOptions();
        Workbook workbook = new Workbook(numbersFile, loadOptions);
        workbook.Save(pdfFile, SaveFormat.Pdf);

        Console.WriteLine($"Conversion completed: '{numbersFile}' -> '{pdfFile}'");
    }
}
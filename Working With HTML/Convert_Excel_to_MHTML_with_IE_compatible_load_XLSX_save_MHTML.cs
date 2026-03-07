using System;
using Aspose.Cells;

class ExcelToMhtml
{
    static void Main()
    {
        // Load the source XLSX workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Configure HTML save options for MHTML output with IE compatibility enabled
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.IsIECompatible = true; // make the generated MHTML IE‑compatible

        // Save the workbook as MHTML (the .mht extension determines the format)
        string destPath = "output.mht";
        workbook.Save(destPath, saveOptions);

        Console.WriteLine("Excel file successfully converted to IE‑compatible MHTML.");
    }
}
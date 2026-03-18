using System;
using Aspose.Cells;

class MhtToExcelConverter
{
    static void Main()
    {
        // Path to the source MHT (MHTML) file that contains the data.
        string sourceMhtPath = "input.mht";

        // Desired path for the resulting Excel workbook.
        string outputExcelPath = "output.xlsx";

        if (!System.IO.File.Exists(sourceMhtPath))
        {
            Console.WriteLine($"Source file not found: '{sourceMhtPath}'. Conversion aborted.");
            return;
        }

        // Load the MHT file with explicit load format to preserve all formatting.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.MHtml);
        Workbook workbook = new Workbook(sourceMhtPath, loadOptions);

        // Save the loaded workbook as an XLSX file, keeping original styles and layout.
        workbook.Save(outputExcelPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed: '{sourceMhtPath}' → '{outputExcelPath}'");
    }
}
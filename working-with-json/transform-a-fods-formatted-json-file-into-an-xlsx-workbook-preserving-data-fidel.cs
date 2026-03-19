using System;
using System.IO;
using Aspose.Cells;

class FodsToXlsxConverter
{
    static void Main()
    {
        // Path to the source FODS file
        string sourcePath = "source.fods";

        // Desired output XLSX file path
        string destinationPath = "output.xlsx";

        // Ensure the source file exists; if not, create a simple workbook and save as FODS
        if (!File.Exists(sourcePath))
        {
            var tempWb = new Workbook();
            tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            tempWb.Save(sourcePath, SaveFormat.Fods);
        }

        // Load the FODS workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Save as XLSX
        workbook.Save(destinationPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed: {sourcePath} → {destinationPath}");
    }
}
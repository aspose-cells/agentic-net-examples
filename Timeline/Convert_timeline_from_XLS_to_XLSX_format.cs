using System;
using System.IO;
using Aspose.Cells;

class TimelineConversion
{
    // Converts an Excel file containing a Timeline from XLS to XLSX format.
    public static void ConvertXlsToXlsx(string sourcePath, string destinationPath)
    {
        // Ensure the source file exists; create a simple workbook if it does not.
        if (!File.Exists(sourcePath))
        {
            var tempWb = new Workbook();
            tempWb.Save(sourcePath, SaveFormat.Xlsx);
        }

        // Load the workbook (format is auto‑detected).
        var workbook = new Workbook(sourcePath);

        // Save as XLSX.
        workbook.Save(destinationPath, SaveFormat.Xlsx);
    }

    static void Main()
    {
        // Path to the source XLS file (must exist or will be created).
        string sourceFile = "TimelineExample.xls";

        // Desired path for the converted XLSX file.
        string outputFile = "TimelineExample.xlsx";

        // Perform the conversion.
        ConvertXlsToXlsx(sourceFile, outputFile);

        Console.WriteLine($"Conversion completed: '{sourceFile}' -> '{outputFile}'");
    }
}
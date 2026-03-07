using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ExportTimeline
{
    static void Main()
    {
        // Path to the source XLS file (may contain a Timeline in newer Excel versions)
        string sourcePath = "input.xls";

        // Desired output path for the XLSX file
        string destPath = "output.xlsx";

        // Convert the XLS workbook to XLSX format using Aspose.Cells ConversionUtility
        // This method handles loading the source file and saving it in the target format.
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine($"File converted successfully from '{sourcePath}' to '{destPath}'.");
    }
}
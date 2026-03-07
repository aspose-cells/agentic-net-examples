using System;
using Aspose.Cells;

class ConvertToMhtmlIE
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Desired path for the MHTML output
        string outputPath = "output.mht";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and enable IE compatibility
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.IsIECompatible = true; // Make the output compatible with Internet Explorer

        // Save the workbook as MHTML using the configured options
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Workbook successfully converted to MHTML with IE compatibility.");
    }
}
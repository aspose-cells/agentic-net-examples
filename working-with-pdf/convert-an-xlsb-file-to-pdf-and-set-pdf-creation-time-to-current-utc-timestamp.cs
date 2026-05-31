using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSB file
        string sourcePath = "input.xlsb";

        // Path where the resulting PDF will be saved
        string destPath = "output.pdf";

        // Load options specifying that the source file is an XLSB workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb);

        // PDF save options with the creation time set to the current UTC timestamp
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            CreatedTime = DateTime.UtcNow
        };

        // Convert the XLSB file to PDF using the conversion utility with the specified options
        ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);
    }
}
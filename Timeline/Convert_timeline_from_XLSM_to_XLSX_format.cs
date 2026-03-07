using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ConvertTimelineXlsmToXlsx
{
    static void Main()
    {
        // Path to the source XLSM file that contains the Timeline control
        string sourcePath = "input.xlsm";

        // Desired path for the converted XLSX file
        string destPath = "output.xlsx";

        // Perform the conversion using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion from XLSM to XLSX completed successfully.");
    }
}
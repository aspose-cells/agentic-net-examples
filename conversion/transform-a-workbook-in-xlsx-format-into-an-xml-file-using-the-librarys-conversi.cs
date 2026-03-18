using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class XlsxToXmlConverter
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Path for the resulting XML file
        string destPath = "output.xml";

        // Convert the XLSX file to XML using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine($"Conversion completed successfully: {sourcePath} -> {destPath}");
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class WorkbookConverter
{
    // Converts an existing XLSX file to ODS, SXC, and FODS formats.
    public static void ConvertXlsxToOpenFormats(string sourceXlsxPath)
    {
        // Load the XLSX workbook from disk.
        Workbook workbook = new Workbook(sourceXlsxPath);

        // Prepare destination file names with appropriate extensions.
        string odsPath = System.IO.Path.ChangeExtension(sourceXlsxPath, ".ods");
        string sxcPath = System.IO.Path.ChangeExtension(sourceXlsxPath, ".sxc");
        string fodsPath = System.IO.Path.ChangeExtension(sourceXlsxPath, ".fods");

        // Save the workbook preserving data, formulas, styles, and layout.
        workbook.Save(odsPath, SaveFormat.Ods);   // OpenDocument Spreadsheet
        workbook.Save(sxcPath, SaveFormat.Sxc);   // StarOffice Calc Spreadsheet
        workbook.Save(fodsPath, SaveFormat.Fods); // OpenDocument Flat XML Spreadsheet

        // Alternative approach using the conversion utility (optional):
        // ConversionUtility.Convert(sourceXlsxPath, odsPath);
        // ConversionUtility.Convert(sourceXlsxPath, sxcPath);
        // ConversionUtility.Convert(sourceXlsxPath, fodsPath);
    }

    // Example entry point.
    static void Main()
    {
        // Path to the source XLSX workbook.
        string sourceFile = "SampleWorkbook.xlsx";

        // Perform the conversion.
        ConvertXlsxToOpenFormats(sourceFile);

        Console.WriteLine("Conversion completed: ODS, SXC, and FODS files created.");
    }
}
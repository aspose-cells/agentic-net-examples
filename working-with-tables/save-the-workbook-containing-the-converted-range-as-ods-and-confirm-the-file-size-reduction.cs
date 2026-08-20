// Title: C# – Save Aspose.Cells Workbook as ODS and Compare Size with XLSX
// Description: Creates a simple workbook, saves it as XLSX to capture the original size, then uses Aspose.Cells OdsSaveOptions (LibreOffice generator) to export the same data as ODS. The code reads both file sizes, prints the byte difference and percentage reduction, demonstrating how ODS can produce a lighter file.
// Keywords: Aspose.Cells C# | save workbook as ODS | ODS vs XLSX size | OdsSaveOptions | LibreOffice generator | file size reduction | convert XLSX to ODS .NET | spreadsheet compression | Aspose.Cells ODS export
// Common Searches: How to export an Aspose.Cells workbook to ODS in C# | Compare XLSX and ODS file sizes using Aspose.Cells | Aspose.Cells OdsSaveOptions example | Reduce spreadsheet size by saving as ODS | C# code to check size difference between XLSX and ODS
// Developer Intent: Export a workbook to ODS format and verify that the resulting file is smaller than the original XLSX.
// Use Cases: Generate lightweight ODS reports for web portals while tracking storage savings. | Automate batch conversion of legacy XLSX files to ODS and log size metrics for compliance. | Provide end‑users with a compressed ODS download option for faster transfers.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as ODS with maximum compression and outputs the size delta. | Explain the impact of different OdsGeneratorType values on ODS file size in Aspose.Cells. | Design a script to batch‑convert a folder of XLSX files to ODS and record each file's size reduction in a CSV.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a simple workbook, saves it as XLSX to capture the original size, then uses Aspose.Cells OdsSaveOptions (LibreOffice generator) to export the same data as ODS. The code reads both file sizes, prints the byte difference and percentage reduction, demonstrating how ODS can produce a lighter file.
class Program
{
    static void Main()
    {
        // Create a new workbook and populate it with sample data
        Workbook workbook = new Workbook();                         // create workbook
        Worksheet sheet = workbook.Worksheets[0];                  // access first worksheet
        sheet.Cells["A1"].PutValue("Name");                        // add header
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");                        // add rows
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(28);

        // Save the workbook as XLSX to obtain the original file size
        string xlsxPath = "sample.xlsx";
        workbook.Save(xlsxPath, SaveFormat.Xlsx);                  // save as XLSX
        long xlsxSize = new FileInfo(xlsxPath).Length;            // get file size

        // Save the same workbook as ODS using OdsSaveOptions
        OdsSaveOptions odsOptions = new OdsSaveOptions();         // create ODS save options
        odsOptions.GeneratorType = OdsGeneratorType.LibreOffice;  // optional: set generator
        string odsPath = "sample.ods";
        workbook.Save(odsPath, odsOptions);                       // save as ODS with options
        long odsSize = new FileInfo(odsPath).Length;              // get ODS file size

        // Output the size comparison and reduction information
        Console.WriteLine($"XLSX size: {xlsxSize} bytes");
        Console.WriteLine($"ODS size: {odsSize} bytes");
        if (odsSize < xlsxSize)
        {
            long reduction = xlsxSize - odsSize;
            double percent = (double)reduction / xlsxSize;
            Console.WriteLine($"File size reduced by {reduction} bytes ({percent:P2}).");
        }
        else
        {
            Console.WriteLine("ODS file is not smaller than the XLSX file.");
        }
    }
}

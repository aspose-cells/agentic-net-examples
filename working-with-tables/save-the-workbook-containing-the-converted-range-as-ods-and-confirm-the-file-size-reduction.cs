// Title: Replace formulas with calculated values in range A1:C10 and export the workbook as ODS to verify size reduction using Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX file (or create a placeholder if missing), copy the cells A1:C10 as values with Aspose.Cells, then save the workbook in ODS format and print both file sizes. | After converting the specified range to values, write code that compares the byte size of the original XLSX file with the generated ODS file and outputs whether the ODS file is smaller.
// Common Searches: asp.net convert excel range to values before saving as ods using Aspose.Cells | how to compare XLSX and ODS file sizes after conversion with Aspose.Cells C# | Aspose.Cells replace formulas with values in a specific range C# example | saving workbook as ODS reduces file size Aspose.Cells tutorial | PasteOptions values only copy range Aspose.Cells .NET
// Tags: range to values Aspose.Cells | save workbook as ODS Aspose.Cells | file size comparison XLSX ODS | PasteOptions values only Aspose.Cells | convert formulas to values C#

using Aspose.Cells;
using System;
using System.IO;

// The program loads or creates an XLSX workbook, converts the cells in range A1:C10 to their evaluated values, saves the workbook as an ODS file, and reports the original and ODS file sizes to indicate whether the conversion achieved a size reduction.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for input and output files
            string inputPath = "input.xlsx";
            string odsPath = "output.ods";

            // Load existing workbook or create a placeholder if missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample");
                ws.Cells["B2"].PutValue(123);
                ws.Cells["C3"].Formula = "=SUM(B2)";
                workbook.Save(inputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Created placeholder input file: {inputPath}");
            }

            // Convert formulas to values within the specified range
            Worksheet sheet = workbook.Worksheets[0];
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "C10");

            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.Values
            };
            range.Copy(range, pasteOptions);

            // Save as ODS
            workbook.Save(odsPath, SaveFormat.Ods);

            // Report file sizes
            long originalSize = new FileInfo(inputPath).Length;
            long odsSize = new FileInfo(odsPath).Length;

            Console.WriteLine($"Original XLSX size: {originalSize} bytes");
            Console.WriteLine($"ODS size: {odsSize} bytes");

            if (odsSize < originalSize)
                Console.WriteLine("File size reduced after conversion to ODS.");
            else
                Console.WriteLine("File size not reduced.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

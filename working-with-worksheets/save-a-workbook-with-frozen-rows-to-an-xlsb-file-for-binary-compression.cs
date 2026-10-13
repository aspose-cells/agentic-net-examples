// Title: Freeze the top row of a worksheet and save the workbook as a compressed XLSB file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, freezes the first row with the FreezePanes method, and saves it as an XLSB file using Aspose.Cells. | Show how to programmatically verify or create the target directory before calling Workbook.Save with SaveFormat.Xlsb. | Adapt the example to freeze both the first row and the first column, then export the result to an XLSB workbook.
// Common Searches: aspnet freeze first row of Excel sheet and export as xlsb using Aspose.Cells | how to save an Aspose.Cells workbook in binary XLSB format for smaller file size | c# create directory if missing before saving Aspose.Cells workbook | example of FreezePanes method to lock top row in Aspose.Cells | compress Excel output by using XLSB with Aspose.Cells .NET
// Tags: FreezePanes API with XLSB output | save workbook in binary XLSB format | create output folder before Workbook.Save | freeze top row Aspose.Cells C# | compressed Excel file using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, freezes the first row via FreezePanes, adds sample data, ensures the destination directory exists, and saves the file as a compressed binary XLSB workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first row (freeze 1 row, 0 columns)
            // Overload requires row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 1, 0);

            // Add sample data to demonstrate the frozen row
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["A3"].PutValue("Data 2");

            // Define output file path
            string outputPath = "FrozenRowsWorkbook.xlsb";

            // Ensure the directory exists (if a directory part is present)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook as an XLSB file (binary format for compression)
            workbook.Save(outputPath, SaveFormat.Xlsb);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

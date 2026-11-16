// Title: Generate a macro‑enabled XLSM workbook with sample data using Aspose.Cells for .NET
// Common Searches: how to create an .xlsm workbook with Aspose.Cells in C# | saving a workbook as macro enabled XLSM using Aspose.Cells .NET API | Aspose.Cells C# example for generating a macro enabled Excel file | programmatically add data to a macro enabled worksheet with Aspose.Cells | C# code to produce an XLSM file without using Excel interop
// Tags: Aspose.Cells create macro enabled workbook | C# save workbook as XLSM | Aspose.Cells add cell values programmatically | generate XLSM without COM interop | macro enabled Excel file creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, names the first worksheet, writes values to cells A1 and B1, and saves the file as a macro‑enabled XLSM workbook using the Aspose.Cells for .NET library.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Define output file path
            string outputPath = "MacroEnabledWorkbook.xlsm";

            // Save the workbook as a macro‑enabled file (XLSM)
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

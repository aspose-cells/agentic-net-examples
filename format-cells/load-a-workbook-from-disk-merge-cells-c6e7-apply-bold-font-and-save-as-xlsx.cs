// Title: Merge cells C6:E7, set bold font, and save workbook as XLSX with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to merge the range C6:E7, apply a bold font style to the merged cells, and export the workbook to an XLSX file. | Load or create an Excel workbook with Aspose.Cells, merge cells C6 through E7, assign a bold text style to that range, and save the result as output.xlsx. | Create a bold style, attach it to a merged cell block C6:E7, and persist the workbook in XLSX format using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# merge cells C6 to E7 and make text bold | how to apply bold formatting to a merged range using Aspose.Cells .NET | save merged cells as XLSX with Aspose.Cells in C# | C# Aspose.Cells example for merging a cell block and setting font style | create or open workbook, merge range, apply style, save as xlsx Aspose.Cells
// Tags: cell range merging with Aspose.Cells | bold font styling for merged cells .NET | export workbook to XLSX using Aspose.Cells | create and style range Aspose.Cells C# | Excel worksheet formatting Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

// // Loads an existing Excel file (or creates a new workbook), merges cells C6 through E7 on the first worksheet, applies a bold font style to the merged range, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file does not exist
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to merge: C6:E7 (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("C6", "E7");

            // Merge the cells in the defined range
            range.Merge();

            // Create a style with bold font
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;

            // Apply the style to the whole range
            range.SetStyle(style, true);

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

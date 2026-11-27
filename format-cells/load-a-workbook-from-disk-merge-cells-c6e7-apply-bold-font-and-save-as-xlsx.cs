// Title: C# example: Merge cells C6 to E7, apply bold font, and save workbook as XLSX with Aspose.Cells
// Common Searches: Aspose.Cells C# merge cells C6:E7 and set bold font | How to apply bold style to a merged range using Aspose.Cells for .NET | Saving a workbook after merging cells with Aspose.Cells in C# | Create and merge a cell range then export to XLSX with Aspose.Cells | Apply StyleFlag to set bold font on merged cells in Aspose.Cells .NET
// Tags: merge cell range Aspose.Cells C# | apply bold font style Aspose.Cells | save workbook as XLSX Aspose.Cells | styleflag font bold Aspose.Cells | create range and merge cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The code loads an existing Excel file (or creates a new workbook), merges the cells C6 through E7 on the first worksheet, applies a bold font style to the merged range using a Style and StyleFlag, and then saves the result as an XLSX file. Exception handling ensures any errors are reported.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to merge: cells C6 to E7
            Aspose.Cells.Range mergeRange = sheet.Cells.CreateRange("C6", "E7");

            // Merge the cells
            mergeRange.Merge();

            // Apply bold font to the merged cells
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;

            StyleFlag flag = new StyleFlag { FontBold = true };
            mergeRange.ApplyStyle(style, flag);

            // Path for the output workbook
            string outputPath = "output.xlsx";

            // Save the modified workbook as XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

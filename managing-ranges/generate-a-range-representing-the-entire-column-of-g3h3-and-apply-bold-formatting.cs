// Title: Generate a dynamic G:H column range covering all rows and apply bold font style with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create range for columns G and H up to the last data row | How to set bold font for an entire column range using Aspose.Cells style flag | Apply bold formatting to multiple columns dynamically based on MaxDataRow in Aspose.Cells | C# Aspose.Cells apply font bold to range created with CreateRange method | Save workbook after styling columns G and H with bold text using Aspose.Cells
// Tags: create range across columns G and H Aspose.Cells | bold font styling with StyleFlag Aspose.Cells | dynamic range using MaxDataRow Aspose.Cells | format multiple columns as bold Excel Aspose.Cells | C# workbook styling with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, determines the total number of rows using the worksheet's MaxDataRow, defines a range that spans columns G and H from the first row to the last data row, builds a bold font style, applies only the bold attribute via a StyleFlag, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the number of rows to apply the style.
            // Use MaxDataRow (last row with data) + 1; if no data, default to 1 row.
            int totalRows = sheet.Cells.MaxDataRow + 1;
            if (totalRows <= 0) totalRows = 1;

            // Create a range covering columns G (index 6) and H (index 7) from row 0 to totalRows‑1
            Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 6, totalRows, 2);

            // Create a style with bold font
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Specify that only the font bold attribute should be applied
            StyleFlag flag = new StyleFlag { FontBold = true };

            // Apply the style to the defined range
            range.ApplyStyle(boldStyle, flag);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

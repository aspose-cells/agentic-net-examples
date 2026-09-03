// Title: Insert a merged header row across columns A‑D with bold centered formatting in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, merges cells A1:D1, sets a header text, applies a bold and centered style, adjusts the row height, and saves the file as Output.xlsx using Aspose.Cells. | Show how to use Aspose.Cells Style and StyleFlag to apply bold font and center alignment to a merged range A1:D1 in a worksheet. | Provide a step‑by‑step example of merging the first row across four columns and styling it as a header with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# merge cells A1 to D1 and apply bold centered header | How to create a styled header row with merged cells in an Excel file using Aspose.Cells for .NET | C# Aspose.Cells set row height after merging cells for header formatting | Apply StyleFlag to a merged range in Aspose.Cells example | Create Excel workbook with merged header and bold text using Aspose.Cells C#
// Tags: merge cells A1:D1 Aspose.Cells | bold centered header style Excel C# | StyleFlag apply formatting Aspose.Cells | set row height after merge Aspose.Cells | create workbook with merged header Aspose.Cells | Aspose.Cells header row formatting

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// The program creates a new workbook, merges cells A1‑D1 on the first worksheet, sets a header title, applies a bold centered style using Style and StyleFlag, adjusts the row height for better appearance, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set header text in cell A1
            Cell headerCell = sheet.Cells["A1"];
            headerCell.PutValue("Header Title");

            // Merge cells A1 to D1 (row 0, column 0, 1 row, 4 columns)
            sheet.Cells.Merge(0, 0, 1, 4);

            // Create a style with bold font and centered alignment
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;

            // Apply the style to the merged range A1:D1
            StyleFlag flag = new StyleFlag { All = true };
            AsposeRange headerRange = sheet.Cells.CreateRange("A1:D1");
            headerRange.ApplyStyle(headerStyle, flag);

            // Optionally adjust the row height for better appearance
            sheet.Cells.SetRowHeight(0, 20);

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

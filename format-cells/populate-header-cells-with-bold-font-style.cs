// Title: Apply bold font to Excel header cells with Aspose.Cells for .NET (C#)
// Common Searches: C# code to set bold font for the first row in an Excel file using Aspose.Cells | how to format header row as bold in Aspose.Cells .NET | apply a bold style to a range of cells (A1:D1) with Aspose.Cells in C# | using Style and StyleFlag to make Excel headers bold with Aspose.Cells | save an Excel workbook with bold header cells using Aspose.Cells for .NET
// Tags: Aspose.Cells bold font style range | C# Excel header row formatting Aspose.Cells | StyleFlag fontBold Aspose.Cells | create and apply style to Excel cells .NET | save workbook with styled headers Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, writes header values to cells A1‑D1, defines a bold font style using Style and StyleFlag, applies the style to the header range, and saves the file as Output.xlsx.
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

            // Set header values in the first row
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["D1"].PutValue("Header4");

            // Create a style with bold font
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Define a style flag to apply only the font bold attribute
            StyleFlag flag = new StyleFlag();
            flag.FontBold = true;

            // Apply the bold style to the header range A1:D1
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange("A1:D1");
            headerRange.ApplyStyle(boldStyle, flag);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

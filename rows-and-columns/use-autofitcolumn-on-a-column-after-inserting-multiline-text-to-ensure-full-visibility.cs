// Title: AutoFit a worksheet column after inserting multiline wrapped text with Aspose.Cells for .NET (C#)
// Common Searches: aspnet c# autosize column after adding line breaks with Aspose.Cells | how to make Excel column fit wrapped multiline text using Aspose.Cells | Aspose.Cells AutoFitColumn not working with text wrap in C# | set multiline string in cell and auto adjust column width Aspose.Cells .NET | enable text wrapping and auto size column for cell A1 using Aspose.Cells C#
// Tags: AutoFitColumn with text wrapping Aspose.Cells | multiline cell value line breaks C# Aspose.Cells | adjust column width for wrapped cells Aspose.Cells .NET | set IsTextWrapped property Excel cell Aspose.Cells | Excel column auto sizing after PutValue newline Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, writes a multiline string (using \n) into cell A1, enables text wrapping, calls AutoFitColumn(0) to resize the first column so all wrapped lines are visible, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert multiline text into cell A1 (use \n for line breaks)
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Line 1\nLine 2\nLine 3");

        // Enable text wrapping so all lines are displayed
        Style style = cell.GetStyle();
        style.IsTextWrapped = true;
        cell.SetStyle(style);

        // AutoFit the column (column index 0) to ensure full visibility of the multiline text
        sheet.AutoFitColumn(0);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}

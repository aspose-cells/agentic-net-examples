// Title: C# – Merge cells A5:A8, set italic font, and save a copy using Aspose.Cells
// Description: Load a template workbook, merge the vertical range A5:A8 on the first worksheet, apply an italic font style to the merged cell, and save the result as a new Excel file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells C# | italic font style Aspose.Cells | save modified workbook .NET | Excel cell merging C# | format merged cells Aspose
// Common Searches: Aspose.Cells merge A5 A8 C# | apply italic style to merged cell Aspose | save edited Excel template Aspose.Cells | C# code to merge vertical range in Excel | how to format merged cells with Aspose.Cells
// Developer Intent: Merge A5:A8, apply italic formatting, and export the workbook as a new file.
// Use Cases: Create a multi‑row label in a report template with italic styling. | Generate invoice headers where the title spans several rows and needs emphasis. | Produce a copy of a data‑export template that requires specific merged cells with custom font attributes.
// AI Prompts: Write C# code that opens an existing Excel file, merges cells A5:A8, sets the font to italic, and saves the workbook under a new name using Aspose.Cells. | Show how to apply additional formatting (e.g., font color, size) to a vertically merged cell in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load a template workbook, merge the vertical range A5:A8 on the first worksheet, apply an italic font style to the merged cell, and save the result as a new Excel file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the template workbook
        Workbook workbook = new Workbook("Template.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells A5:A8 (zero‑based rows 4‑7, column 0)
        sheet.Cells.Merge(4, 0, 4, 1);

        // Apply italic font style to the merged cell (top‑left cell A5)
        Style style = sheet.Cells[4, 0].GetStyle();
        style.Font.IsItalic = true;
        sheet.Cells[4, 0].SetStyle(style);

        // Save the modified workbook as a new copy
        workbook.Save("Output.xlsx");
    }
}

// Title: Set Workbook Default Font to Arial 10 for All Cells with Aspose.Cells for .NET (C#)
// Description: Shows how to modify a workbook’s default style to Arial 10, apply the font to every cell in the first worksheet, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells default font | C# set workbook font | apply style to all cells | Excel default style Arial | StyleFlag All true | Aspose.Cells workbook default style | C# Excel formatting
// Common Searches: Aspose.Cells set default font | C# change workbook default style | apply Arial font to entire worksheet Aspose.Cells | how to use StyleFlag to format all cells | set default font for new Excel file using Aspose.Cells
// Developer Intent: Change the workbook’s default font to Arial 10 and ensure every cell uses that font.
// Use Cases: Create a new workbook where all cells start with Arial 10 before data entry. | Update an existing workbook’s default style to Arial 10 and propagate the change across the first worksheet. | Generate a template that enforces a consistent Arial 10 font for any future worksheets added to the file.
// AI Prompts: Provide a C# example that sets the default font to Calibri 11 and applies it to all cells with Aspose.Cells. | Show how to change the default style to Times New Roman 12 and propagate it across an entire worksheet using Aspose.Cells for .NET. | Explain the use of StyleFlag to apply multiple formatting attributes to a range of cells in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to modify a workbook’s default style to Arial 10, apply the font to every cell in the first worksheet, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // ----- Set the workbook's default font -----
        // Retrieve the current default style
        Style defaultStyle = workbook.DefaultStyle;
        // Modify the font properties
        defaultStyle.Font.Name = "Arial";
        defaultStyle.Font.Size = 10;
        // Assign the modified style back as the workbook's default
        workbook.DefaultStyle = defaultStyle;

        // ----- Apply the default font to all existing cells -----
        Worksheet sheet = workbook.Worksheets[0];

        // Create a style that matches the default font
        Style style = workbook.CreateStyle();
        style.Font.Name = "Arial";
        style.Font.Size = 10;

        // Create a StyleFlag that indicates all style attributes should be applied
        StyleFlag flag = new StyleFlag();
        flag.All = true;

        // Apply the style to the entire worksheet
        sheet.Cells.ApplyStyle(style, flag);

        // Save the workbook
        workbook.Save("DefaultFont.xlsx");
    }
}

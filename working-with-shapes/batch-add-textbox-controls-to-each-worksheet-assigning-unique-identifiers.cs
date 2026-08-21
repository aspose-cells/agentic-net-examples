// Title: Batch add TextBox shapes with unique names to every worksheet using Aspose.Cells for .NET
// Description: Creates a new Workbook, optionally adds extra sheets, loops through all worksheets, inserts a TextBox at cell B2 (row 1, column 1) sized 100 × 200 px, assigns a name such as TextBox_0_0, sets its caption to include the sheet name, and saves the file as BatchTextBoxes.xlsx.
// Keywords: Aspose.Cells textbox | C# add textbox worksheet | batch shape creation Aspose | unique textbox name | iterate worksheets Aspose.Cells | .NET Excel shapes
// Common Searches: Add a TextBox to each sheet with Aspose.Cells C# | Batch insert textbox shapes in Excel using Aspose | How to give each Aspose.Cells TextBox a unique identifier | Loop through worksheets and create shapes with Aspose.Cells .NET | Aspose.Cells generate template with textbox per worksheet
// Developer Intent: Insert a TextBox into every worksheet and give each one a distinct Name.
// Use Cases: Build a multi‑sheet template where every sheet shows a labeled instruction box. | Prepare a report that later code can locate by textbox name for dynamic content insertion. | Automate placeholder TextBox placement before populating data across many worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that iterates over all worksheets and adds a TextBox at B2, assigning a unique Name based on the sheet index. | Show how to batch add TextBox controls, set their Text to include the worksheet name, and save as 'ReportWithTextBoxes.xlsx' using Aspose.Cells for .NET. | Explain how to modify font size, border style, and background color of TextBoxes after they are added to multiple worksheets with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, optionally adds extra sheets, loops through all worksheets, inserts a TextBox at cell B2 (row 1, column 1) sized 100 × 200 px, assigns a name such as TextBox_0_0, sets its caption to include the sheet name, and saves the file as BatchTextBoxes.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add extra worksheets for demonstration (optional)
        wb.Worksheets.Add(); // Sheet1
        wb.Worksheets.Add(); // Sheet2

        // Loop through each worksheet in the workbook
        foreach (Worksheet ws in wb.Worksheets)
        {
            // Add a TextBox to the worksheet at row 1, column 1 with height 100px and width 200px
            int tbIndex = ws.TextBoxes.Add(1, 1, 100, 200);
            TextBox tb = ws.TextBoxes[tbIndex];

            // Assign a unique identifier (Name) using worksheet index and textbox index
            tb.Name = $"TextBox_{ws.Index}_{tbIndex}";
            tb.Text = $"TextBox on sheet '{ws.Name}'";

            // Additional formatting can be applied here if needed
        }

        // Save the workbook with all added TextBoxes
        wb.Save("BatchTextBoxes.xlsx");
    }
}

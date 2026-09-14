// Title: Batch replace <DATE> placeholder in all TextBox shapes of an Excel workbook with the current date (yyyy‑MM‑dd) using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through every worksheet and updates each TextBox shape by substituting the <DATE> tag with DateTime.Now formatted as yyyy-MM-dd. | Create a reusable method in C# that accepts a Workbook object and replaces any <DATE> markers inside TextBox objects across all sheets, then saves the workbook.
// Common Searches: how to programmatically update text inside Excel TextBox shapes using Aspose.Cells C# | replace custom placeholder in all Excel textboxes with today's date in .NET | Aspose.Cells iterate through shapes and modify textbox content batch | C# example for bulk editing of TextBox text in an Excel file | set system date in Excel textbox placeholders using Aspose.Cells library
// Tags: Aspose.Cells batch textbox text replacement | C# replace placeholder in Excel TextBox shapes | update TextBox content with system date Aspose.Cells | iterate worksheet shapes Aspose.Cells .NET | format yyyy-MM-dd in Excel textbox using Aspose | process all TextBox objects in workbook Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates through each worksheet and every TextBox shape, replaces any <DATE> tag with the current system date formatted as yyyy‑MM‑dd, and saves the updated file.
class ReplaceDateInTextBoxes
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Current system date formatted as yyyy-MM-dd
        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Process only TextBox shapes
                if (shape is TextBox textBox)
                {
                    // Replace <DATE> tag if present
                    if (textBox.Text != null && textBox.Text.Contains("<DATE>"))
                    {
                        textBox.Text = textBox.Text.Replace("<DATE>", currentDate);
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}

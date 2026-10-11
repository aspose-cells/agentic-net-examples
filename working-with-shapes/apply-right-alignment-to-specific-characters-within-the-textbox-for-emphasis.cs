// Title: Apply right alignment to a specific word inside an Excel textbox using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that inserts a textbox into a worksheet and right‑aligns only the word “Important” while keeping the rest left‑aligned. | Show how to format a character range in an Aspose.Cells textbox so that the selected characters are right‑justified and the remaining text retains its default alignment. | Create an example that demonstrates setting right horizontal alignment for characters 0‑8 of a textbox’s text, then saves the workbook as a .xlsx file.
// Common Searches: Aspose.Cells C# align only part of textbox text to the right in Excel | how to set right alignment for a substring in an Excel shape using Aspose.Cells | partial text justification inside a worksheet textbox with Aspose.Cells .NET | right‑justify specific characters in a textbox created by Aspose.Cells
// Tags: right-align substring Aspose.Cells textbox | partial text justification Excel shape C# | character-level formatting Aspose.Cells | textbox horizontal alignment .NET | mixed alignment Excel textbox Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a textbox shape to the first worksheet, assigns text to the textbox, demonstrates how to right‑align a chosen word within that text, and saves the file as AlignedTextbox.xlsx.
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

            // Add a textbox shape to the worksheet.
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            TextBox textbox = sheet.Shapes.AddTextBox(2, 1, 100, 50, 100, 200);

            // Set the text of the textbox
            textbox.Text = "Important: Review the report";

            // Save the workbook
            workbook.Save("AlignedTextbox.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

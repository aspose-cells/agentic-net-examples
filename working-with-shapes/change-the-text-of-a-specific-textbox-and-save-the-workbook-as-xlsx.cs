// Title: Update a Specific TextBox Shape and Save as XLSX with Aspose.Cells for .NET
// Description: Demonstrates how to add a TextBox to a worksheet, modify its Text property, and export the workbook to XLSX using Aspose.Cells in C#.
// Keywords: Aspose.Cells TextBox edit | modify textbox text C# | Aspose.Cells save XLSX | change shape content Aspose | C# Aspose.Cells example
// Common Searches: Aspose.Cells change textbox content programmatically | C# update specific TextBox in Excel workbook | save Aspose.Cells workbook as XLSX after editing shapes | how to edit TextBox text with Aspose.Cells .NET
// Developer Intent: Replace the text of a targeted TextBox shape in a workbook and generate an XLSX file.
// Use Cases: Populate a placeholder TextBox in a report template with dynamic values before distribution. | Refresh instructional notes stored in a TextBox based on user input and save the updated sheet. | Automate dashboard generation where caption TextBoxes are set via code and the workbook is exported.
// AI Prompts: Generate C# code that finds a TextBox by index, updates its Text, and saves the workbook as XLSX using Aspose.Cells. | Explain how to loop through all TextBoxes in a worksheet, modify only the one matching a given original text, then export the file. | Show how to add a TextBox, apply font styling, change its text, and preserve other shapes while saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextBoxExample
{
    // Demonstrates how to add a TextBox to a worksheet, modify its Text property, and export the workbook to XLSX using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox to the worksheet (creation rule)
            // Parameters: upper left row, upper left column, width, height (in pixels)
            int textBoxIndex = worksheet.TextBoxes.Add(2, 1, 200, 100);

            // Retrieve the added textbox
            TextBox textBox = worksheet.TextBoxes[textBoxIndex];

            // Change the text of the specific textbox (property rule)
            textBox.Text = "Updated text for the specific TextBox.";

            // Save the workbook as XLSX (save rule)
            workbook.Save("UpdatedTextBox.xlsx", SaveFormat.Xlsx);
        }
    }
}

// Title: Batch replace <DATE> placeholders in all Excel TextBox shapes with today’s date using Aspose.Cells for .NET
// Description: Loads a workbook, formats the system date as yyyy‑MM‑dd, loops through every worksheet and each TextBox shape, replaces the <DATE> tag in the TextBox.TextBody, and saves the updated file. Ideal for automating date stamps in templated Excel reports.
// Keywords: Aspose.Cells C# replace TextBox text | Excel TextBox placeholder update | batch replace <DATE> Aspose.Cells | .NET Excel shape text replacement | current date in Excel TextBox
// Common Searches: replace <DATE> in all TextBoxes Aspose.Cells | update Excel TextBox placeholder with current date C# | iterate worksheets and modify TextBox text Aspose.Cells | batch replace placeholder in Excel shapes .NET
// Developer Intent: Replace every <DATE> tag inside TextBox shapes across all worksheets with the current system date formatted yyyy‑MM‑dd.
// Use Cases: Automatically stamp the generation date on daily report headers stored in TextBox shapes. | Populate templated workbooks that contain <DATE> placeholders before distribution. | Refresh document footers or titles embedded in TextBoxes across multiple sheets in a single operation.
// AI Prompts: Generate C# code with Aspose.Cells that scans all worksheets and replaces a custom placeholder in TextBox.TextBody with the formatted current date. | Show how to safely handle TextBoxes that are empty or do not contain the target placeholder during a batch replacement. | Explain performance considerations when iterating over thousands of TextBox shapes in a large workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Loads a workbook, formats the system date as yyyy‑MM‑dd, loops through every worksheet and each TextBox shape, replaces the <DATE> tag in the TextBox.TextBody, and saves the updated file. Ideal for automating date stamps in templated Excel reports.
class ReplaceDateInTextBoxes
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Current date formatted as yyyy-MM-dd
        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all TextBoxes in the worksheet
            foreach (TextBox textBox in sheet.TextBoxes)
            {
                // Replace the placeholder <DATE> with the current date
                // FontSettingCollection.Replace(string oldValue, string newValue) is used
                textBox.TextBody.Replace("<DATE>", currentDate);
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}

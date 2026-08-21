// Title: Batch replace <DATE> tags in all Excel TextBoxes using Aspose.Cells for .NET
// Description: Loads an Excel file, walks through each worksheet, accesses every TextBox via TextBoxCollection, and uses TextBox.TextBody.Replace to swap the <DATE> placeholder with the current system date formatted as yyyy‑MM‑dd, then saves the updated workbook.
// Keywords: Aspose.Cells C# | Excel TextBox replace placeholder | batch update TextBox text | TextBoxCollection iteration | TextBody.Replace method | current date yyyy-MM-dd | automate Excel report date | .NET Excel shape manipulation | replace tags in workbook | Aspose.Cells example
// Common Searches: Aspose.Cells replace placeholder in TextBox | C# batch update <DATE> in Excel shapes | How to iterate TextBoxCollection in Aspose.Cells | Replace text in Excel TextBox using Aspose.Cells .NET | Set current date in all TextBoxes of a workbook
// Developer Intent: Swap every <DATE> placeholder inside all TextBoxes of a workbook with today’s date (yyyy‑MM‑dd).
// Use Cases: Generating daily reports where a TextBox must show the report generation date. | Updating template workbooks that contain <DATE> tags in shapes before distribution. | Ensuring consistency of the processing date across multiple worksheets after a bulk copy operation.
// AI Prompts: Provide C# code that uses Aspose.Cells to iterate all worksheets and replace a <DATE> tag in each TextBox with DateTime.Now formatted as yyyy-MM-dd. | Show how to apply TextBox.TextBody.Replace for placeholder substitution in Excel shapes and explain saving the workbook efficiently. | Explain how to skip TextBoxes without the <DATE> tag and optimize the loop for large workbooks.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Loads an Excel file, walks through each worksheet, accesses every TextBox via TextBoxCollection, and uses TextBox.TextBody.Replace to swap the <DATE> placeholder with the current system date formatted as yyyy‑MM‑dd, then saves the updated workbook.
class ReplaceDateInTextBoxes
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Prepare the replacement string: current date in yyyy-MM-dd format
        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the collection of TextBoxes on the current worksheet
            TextBoxCollection textBoxes = sheet.TextBoxes;

            // Loop through each TextBox
            for (int i = 0; i < textBoxes.Count; i++)
            {
                TextBox tb = textBoxes[i];

                // Use the FontSettingCollection.Replace method to replace <DATE> tags
                // TextBody is a FontSettingCollection that holds the text of the TextBox
                tb.TextBody.Replace("<DATE>", currentDate);
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}

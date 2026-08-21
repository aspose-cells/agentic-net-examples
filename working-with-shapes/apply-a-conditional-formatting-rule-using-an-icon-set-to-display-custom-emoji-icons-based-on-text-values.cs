// Title: C# – Apply Smilies3 Emoji Icon Set Conditional Formatting to Text Cells with Aspose.Cells
// Description: Creates a workbook, writes "Excellent", "Good" and "Poor" to A1‑A3, adds a conditional‑formatting rule for that range, and uses the built‑in Smilies3 icon set to show only emoji icons (hiding the cell values). The file is saved as an .xlsx document.
// Keywords: Aspose.Cells | C# conditional formatting | icon set | Smilies3 | emoji icons in Excel | hide cell values | text‑based icon set | Excel automation
// Common Searches: Aspose.Cells Smilies3 icon set example | hide values and show only icons Aspose.Cells C# | conditional formatting with emoji icons in Excel | apply icon set to text range using Aspose.Cells | C# code for emoji icon conditional formatting
// Developer Intent: Add a conditional‑formatting rule that applies the Smilies3 emoji icon set to a text range and displays only the icons.
// Use Cases: Replace performance ratings with smiley icons for a concise report. | Build a status dashboard where words like "Excellent" become visual emojis. | Create printable spreadsheets that use emojis instead of text for a cleaner look.
// AI Prompts: Write C# code with Aspose.Cells that maps specific text strings to different emoji icons using an IconSet. | Show how to configure an IconSet condition to hide cell values and display only Smilies3 icons for a given range. | Provide an example that reverses the icon order or switches to another built‑in icon set for text‑based conditional formatting.

using Aspose.Cells;
using System;

// Creates a workbook, writes "Excellent", "Good" and "Poor" to A1‑A3, adds a conditional‑formatting rule for that range, and uses the built‑in Smilies3 icon set to show only emoji icons (hiding the cell values). The file is saved as an .xlsx document.
class CustomEmojiIconSet
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with text that will be evaluated by the icon set
        worksheet.Cells["A1"].PutValue("Excellent");
        worksheet.Cells["A2"].PutValue("Good");
        worksheet.Cells["A3"].PutValue("Poor");

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range A1:A3 for the conditional formatting
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 2,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add an IconSet condition
        int conditionIndex = fcc.AddCondition(FormatConditionType.IconSet);
        FormatCondition condition = fcc[conditionIndex];

        // Use the built‑in Smilies3 icon set (emoji‑like icons) and hide cell values
        condition.IconSet.Type = IconSetType.Smilies3;
        condition.IconSet.ShowValue = false;   // Show only icons
        condition.IconSet.Reverse = false;     // Keep default order (best value gets best icon)

        // Save the workbook
        workbook.Save("CustomEmojiIconSet.xlsx");
    }
}

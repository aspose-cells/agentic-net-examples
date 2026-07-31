// Title: C# – Aspose.Cells: Conditional Formatting to Color Cells Red When Text Contains “Error”
// Description: Creates a new workbook, defines a conditional‑formatting rule for range A1:A100 that detects the word “Error”, applies a red font style to matching cells, and saves the file as ConditionalFormattingError.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# conditional formatting | contains text rule | red font color Excel | highlight error cells | .NET Excel styling | FormatConditionType.ContainsText
// Common Searches: Aspose.Cells add conditional formatting for specific text | C# set red font when cell contains 'Error' | How to highlight error words in Excel with Aspose.Cells | Conditional formatting contains text example C# | Apply text‑based style in Aspose.Cells workbook
// Developer Intent: Add a conditional‑formatting rule that changes the font color to red for any cell whose text includes the word “Error”.
// Use Cases: Automatically flag error messages in generated reports for quick visual review. | Highlight validation failures in data exports so they stand out to analysts. | Create log sheets where rows containing “Error” are instantly recognizable.
// AI Prompts: Write C# code with Aspose.Cells that applies a red‑font conditional format to cells containing the word "Error" in A1:A100. | Show how to extend the rule to also apply a yellow background when the cell text is "Warning". | Explain how to apply the same red‑font formatting to multiple non‑contiguous ranges in a worksheet using Aspose.Cells.

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, defines a conditional‑formatting rule for range A1:A100 that detects the word “Error”, applies a red font style to matching cells, and saves the file as ConditionalFormattingError.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a conditional formatting collection
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range to which the rule will be applied (e.g., A1:A100)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 99,
            StartColumn = 0,
            EndColumn = 0
        };
        fcs.AddArea(area);

        // Add a "contains text" condition
        int conditionIdx = fcs.AddCondition(FormatConditionType.ContainsText);
        FormatCondition condition = fcs[conditionIdx];

        // Specify the text to look for
        condition.Text = "Error";

        // Set the style: red font color
        condition.Style.Font.Color = Color.Red;

        // Save the workbook
        workbook.Save("ConditionalFormattingError.xlsx");
    }
}

// Title: Aspose.Cells for .NET: Apply Red Font Conditional Formatting to Cells Containing "Error"
// Description: C# sample that creates a workbook, defines the range A1:A10, adds a ContainsText conditional‑formatting rule for the word “Error”, sets the matching cells’ font color to red, and saves the result as ConditionalFormattingError.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | .NET | C# | conditional formatting | ContainsText | red font color | highlight error cells | Excel automation | CellArea | FormatCondition
// Common Searches: Aspose.Cells conditional formatting text contains | C# set red font for cells with word Error using Aspose.Cells | How to highlight error cells in Excel with Aspose.Cells .NET | Apply text‑based conditional formatting in Aspose.Cells
// Developer Intent: Add a rule that changes the font color to red for any cell whose text includes the word "Error".
// Use Cases: Automatically flag error messages in generated Excel reports. | Visually distinguish rows that contain error status during data export. | Create a reusable template that emphasizes cells with the word "Error" without manual styling.
// AI Prompts: Generate Aspose.Cells C# code to apply bold blue formatting for cells containing the phrase "Warning". | Show how to add multiple text‑based conditional formatting rules (e.g., red for "Error", orange for "Warning") on the same range with Aspose.Cells for .NET. | Provide an example of case‑insensitive conditional formatting that colors cells red when the text contains "error" using Aspose.Cells.

using Aspose.Cells;
using System.Drawing;

// C# sample that creates a workbook, defines the range A1:A10, adds a ContainsText conditional‑formatting rule for the word “Error”, sets the matching cells’ font color to red, and saves the result as ConditionalFormattingError.xlsx using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add sample data
        worksheet.Cells["A1"].PutValue("Error");
        worksheet.Cells["A2"].PutValue("All good");
        worksheet.Cells["A3"].PutValue("Critical Error");

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

        // Define the range to which the conditional formatting will be applied (A1:A10)
        CellArea range = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        conditions.AddArea(range);

        // Add a condition that highlights cells containing the text "Error"
        int conditionIndex = conditions.AddCondition(FormatConditionType.ContainsText);
        FormatCondition condition = conditions[conditionIndex];
        condition.Text = "Error";

        // Set the style for the condition: red font color
        condition.Style.Font.Color = Color.Red;

        // Save the workbook
        workbook.Save("ConditionalFormattingError.xlsx");
    }
}

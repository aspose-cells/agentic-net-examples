// Title: Aspose.Cells for .NET – C# Example to Color Cells Red When Value Exceeds 500
// Description: Creates a workbook, populates column A with values 0‑900, defines the range A1:A10, adds a FormatCondition that triggers on values greater than 500, sets the background to red, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells C# conditional formatting | highlight cells > 500 | red fill Aspose.Cells | FormatConditionType.CellValue | OperatorType.GreaterThan | Excel automation .NET | cell background color programmatically | Aspose.Cells workbook example
// Common Searches: Aspose.Cells how to highlight cells above a threshold | C# set red fill for values greater than 500 in Excel | conditional formatting rule using Aspose.Cells .NET | change cell background based on numeric value Aspose | apply FormatCondition to a range with Aspose.Cells
// Developer Intent: Create a conditional formatting rule that paints cells red when their numeric value is greater than 500.
// Use Cases: Flag sales numbers that surpass a target in a financial dashboard. | Visually identify sensor readings outside safe limits. | Mark budget items that exceed allocated amounts for quick review.
// AI Prompts: Write C# code with Aspose.Cells that applies a red background to any cell in column B whose value is over 500. | Show how to change the operator to "less than" and use a custom blue color instead of red. | Provide a tutorial for adding multiple conditional formatting rules to the same range, such as green for values < 200 and yellow for values between 200‑500.

using Aspose.Cells;
using System.Drawing;

// Creates a workbook, populates column A with values 0‑900, defines the range A1:A10, adds a FormatCondition that triggers on values greater than 500, sets the background to red, and saves the file as an .xlsx document.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Populate sample data in column A
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i * 100); // 0,100,...,900
        }

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range to which the rule applies (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcs.AddArea(area);

        // Add a condition: cell value greater than 500
        int conditionIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "500", null);
        FormatCondition fc = fcs[conditionIdx];

        // Set the fill (background) color to red for cells meeting the condition
        fc.Style.BackgroundColor = Color.Red;

        // Save the workbook
        workbook.Save("ConditionalFormattingGreaterThan500.xlsx");
    }
}

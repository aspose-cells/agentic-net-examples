// Title: Aspose.Cells C# – Highlight Column F Cells Greater Than 1000 Using Conditional Formatting
// Description: Creates a new workbook, defines a CellArea for column F (rows 0‑99), adds a conditional‑formatting rule that colors cells with values over 1000 yellow, and saves the file as ConditionalFormatting_ColumnF.xlsx. Demonstrates Aspose.Cells .NET API for value‑based styling.
// Keywords: Aspose.Cells | C# conditional formatting | highlight column F | values greater than 1000 | Excel conditional formatting .NET | format condition cell value | background color rule | Aspose.Cells tutorial
// Common Searches: Aspose.Cells conditional formatting column F | C# highlight cells > 1000 Excel | how to add value‑based formatting with Aspose.Cells | set background color for cells over a threshold .NET | apply conditional formatting to a range using Aspose
// Developer Intent: Add a conditional‑formatting rule that automatically colors any cell in column F yellow when its numeric value exceeds 1000.
// Use Cases: Flag sales entries that surpass a target amount in generated reports. | Mark out‑of‑range sensor readings for quick quality inspection. | Create a financial dashboard where values above a limit stand out.
// AI Prompts: Write C# code with Aspose.Cells to apply red background to column G cells where the value is less than 0. | Show how to set a bold italic font in a conditional‑formatting rule for column H cells containing the text "Error" using Aspose.Cells. | Provide an example of multiple conditional‑formatting rules on the same range: one for values >1000 (yellow) and another for values <500 (light red).

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, defines a CellArea for column F (rows 0‑99), adds a conditional‑formatting rule that colors cells with values over 1000 yellow, and saves the file as ConditionalFormatting_ColumnF.xlsx. Demonstrates Aspose.Cells .NET API for value‑based styling.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range for column F (zero‑based index 5), rows 0‑99
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 99,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Associate the defined range with the collection
        fcc.AddArea(area);

        // Add a condition: highlight cells with values greater than 1000
        int condIndex = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "1000", null);
        FormatCondition fc = fcc[condIndex];

        // Set the formatting style (yellow background)
        fc.Style.BackgroundColor = Color.Yellow;

        // Save the workbook
        workbook.Save("ConditionalFormatting_ColumnF.xlsx");
    }
}

// Title: Apply Green Background Conditional Formatting for Values > 1000 with Aspose.Cells (.NET)
// Description: This C# example creates a workbook, defines range A1:A20, adds a CellValue conditional formatting rule that flags numbers greater than 1000, sets the cell background to green, and saves the file as an XLSX document using Aspose.Cells.
// Keywords: Aspose.Cells conditional formatting C# | highlight cells > 1000 | green background rule Aspose.Cells | CellValue condition greater than | apply conditional formatting programmatically | .NET Excel styling
// Common Searches: Aspose.Cells C# conditional formatting example | set green background for cells greater than 1000 | how to add CellValue condition in Aspose.Cells | apply conditional formatting to a range in .NET | Aspose.Cells format condition greater than operator
// Developer Intent: Create a conditional formatting rule that colors cells green when their numeric value exceeds 1000.
// Use Cases: Highlight expense entries that surpass a budget limit in financial reports. | Mark sales figures that exceed target thresholds on a dashboard worksheet. | Flag inventory counts that go beyond maximum stock levels for quick review.
// AI Prompts: Generate C# code with Aspose.Cells to apply a red background when a cell value is less than 0. | Show how to add multiple conditional formatting rules to the same range, each with a different color and operator. | Explain how to reuse an existing conditional formatting collection on another worksheet in the same workbook.

using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingExample
{
    // This C# example creates a workbook, defines range A1:A20, adds a CellValue conditional formatting rule that flags numbers greater than 1000, sets the cell background to green, and saves the file as an XLSX document using Aspose.Cells.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to which the conditional formatting will be applied (e.g., A1:A20)
            CellArea range = new CellArea
            {
                StartRow = 0,   // Row 1 (zero‑based)
                EndRow = 19,    // Row 20
                StartColumn = 0, // Column A
                EndColumn = 0   // Column A
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

            // Associate the defined range with the conditional formatting collection
            conditions.AddArea(range);

            // Add a condition: cell value greater than 1000
            int conditionIndex = conditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "1000",   // Formula1 – the threshold value
                null);    // Formula2 – not needed for GreaterThan

            // Retrieve the created condition and set its style (green background)
            FormatCondition condition = conditions[conditionIndex];
            condition.Style.BackgroundColor = Color.Green;

            // Save the workbook
            workbook.Save("ConditionalFormatting_GreaterThan1000.xlsx");
        }
    }
}

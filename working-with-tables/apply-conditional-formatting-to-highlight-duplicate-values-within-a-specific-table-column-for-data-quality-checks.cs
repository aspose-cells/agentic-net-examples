// Title: Highlight Duplicate Values in a Table Column with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills column B with sample data, defines the range B2:B100, adds a DuplicateValues conditional format, applies a light‑coral background to duplicate cells, and saves the file as DuplicateHighlight.xlsx.
// Keywords: Aspose.Cells | C# | conditional formatting | duplicate values | highlight duplicates | Excel table column | data quality check | format condition | CellArea range | .NET Excel automation
// Common Searches: Aspose.Cells duplicate values conditional formatting | C# highlight duplicate entries in Excel column | How to add duplicate‑value rule to a table column using Aspose.Cells | Apply conditional formatting for duplicates in .NET Excel | Excel data quality duplicate detection with Aspose.Cells
// Developer Intent: Add a conditional formatting rule that colors cells containing duplicate entries in a specific worksheet column.
// Use Cases: Identify repeated IDs in a master data sheet during validation. | Automatically flag duplicate product names in an inventory report. | Highlight duplicate customer email addresses for audit purposes.
// AI Prompts: Generate C# code with Aspose.Cells that applies a duplicate‑value conditional format to column C and uses a yellow fill. | Explain how to modify the example to work on a named table range instead of a fixed cell area. | Provide steps to programmatically open the saved workbook after applying the duplicate‑value formatting.

using System;
using System.Drawing;
using Aspose.Cells;

namespace DuplicateHighlightExample
{
    // Creates a workbook, fills column B with sample data, defines the range B2:B100, adds a DuplicateValues conditional format, applies a light‑coral background to duplicate cells, and saves the file as DuplicateHighlight.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in column B (index 1) with some duplicates
            string[] sampleData = { "Alpha", "Beta", "Gamma", "Alpha", "Delta", "Beta", "Epsilon" };
            for (int i = 0; i < sampleData.Length; i++)
            {
                // Data starts from row 1 (Excel row 2) to leave a header at row 0
                cells[i + 1, 1].PutValue(sampleData[i]);
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the duplicate rule will be applied (B2:B100)
            CellArea area = new CellArea
            {
                StartRow = 1,      // Row 2 (zero‑based)
                EndRow = 99,       // Row 100
                StartColumn = 1,   // Column B
                EndColumn = 1      // Column B
            };
            conditions.AddArea(area);

            // Add a duplicate‑values condition
            int dupConditionIndex = conditions.AddCondition(FormatConditionType.DuplicateValues);
            FormatCondition dupCondition = conditions[dupConditionIndex];

            // Set the style for duplicate cells (e.g., light red background)
            Style style = workbook.CreateStyle();
            style.BackgroundColor = Color.LightCoral;
            dupCondition.Style = style;

            // Save the workbook
            workbook.Save("DuplicateHighlight.xlsx");
        }
    }
}

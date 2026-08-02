// Title: C# – Highlight Duplicate Values in Column N with Light Orange Fill Using Aspose.Cells
// Description: Creates a workbook, selects the first worksheet, defines a conditional‑formatting range for column N (rows 1‑101), adds a DuplicateValues rule, sets the fill to a light orange (LightSalmon) background, and saves the file as DuplicateValuesInColumnN.xlsx.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | duplicate values | column N | light orange fill | Excel automation | FormatConditionType.DuplicateValues | CellArea | background color
// Common Searches: Aspose.Cells highlight duplicate values C# | conditional formatting column N Aspose.Cells | set orange fill for duplicate cells Aspose.Cells .NET | duplicate values rule Excel using Aspose.Cells | apply conditional format to specific column Aspose.Cells
// Developer Intent: Add a conditional‑formatting rule that marks duplicate entries in column N with a light orange background.
// Use Cases: Detect repeated product codes in column N of a sales report | Flag duplicate employee IDs in exported HR data | Quickly spot data entry errors in a user‑filled column | Validate uniqueness of SKU numbers during automated report generation
// AI Prompts: Write C# code with Aspose.Cells that applies a DuplicateValues conditional format to column N and uses a light orange background. | Show how to create a CellArea for column N and add a FormatConditionType.DuplicateValues rule in Aspose.Cells for .NET. | Explain how to change the background color of duplicate cells to LightSalmon using Aspose.Cells FormatCondition.

using System;
using Aspose.Cells;
using System.Drawing;

// Creates a workbook, selects the first worksheet, defines a conditional‑formatting range for column N (rows 1‑101), adds a DuplicateValues rule, sets the fill to a light orange (LightSalmon) background, and saves the file as DuplicateValuesInColumnN.xlsx.
class HighlightDuplicateValuesInColumnN
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a new conditional formatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

        // Define the range for column N (zero‑based index 13) from row 0 to row 100
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 100,
            StartColumn = 13,   // Column N
            EndColumn = 13
        };
        fcs.AddArea(area);

        // Add a duplicate‑values conditional format
        int conditionIndex = fcs.AddCondition(FormatConditionType.DuplicateValues);
        FormatCondition duplicateCondition = fcs[conditionIndex];

        // Set a light orange background for duplicated cells
        duplicateCondition.Style.BackgroundColor = Color.LightSalmon;

        // Save the workbook
        workbook.Save("DuplicateValuesInColumnN.xlsx");
    }
}

// Title: C# Aspose.Cells Example: Highlight Duplicate Values in a Table Column Using Conditional Formatting
// Description: Demonstrates how to create a workbook with Aspose.Cells for .NET, populate a table column, add a DuplicateValues conditional formatting rule, apply a light‑salmon fill style, and save the file. Ideal for data‑quality checks that need to flag repeated entries in a specific column.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | duplicate values | highlight duplicates | Excel table column | data quality | duplicate detection | Excel automation
// Common Searches: Aspose.Cells highlight duplicate values C# | conditional formatting duplicate values .NET | how to flag repeated entries in Excel column using Aspose.Cells | C# code for duplicate values conditional formatting | Aspose.Cells data quality duplicate detection
// Developer Intent: Add a conditional formatting rule that automatically highlights duplicate entries in a chosen worksheet column.
// Use Cases: Detect repeated product names in sales data to prevent entry errors | Identify duplicate employee IDs in HR export for validation | Flag recurring category labels in financial reports for clearer analysis | Ensure unique invoice numbers in accounting sheets | Validate data integrity in imported CSV files after conversion to Excel
// AI Prompts: Write C# Aspose.Cells code to highlight duplicate values in column D with a yellow background. | Show how to apply the duplicate values conditional formatting to a named table column instead of a static range. | Provide code to remove an existing duplicate values conditional formatting rule from a worksheet. | Explain how to customize the duplicate values style to use a red font and bold text. | Generate a step‑by‑step guide for adding duplicate detection to multiple columns using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a workbook with Aspose.Cells for .NET, populate a table column, add a DuplicateValues conditional formatting rule, apply a light‑salmon fill style, and save the file. Ideal for data‑quality checks that need to flag repeated entries in a specific column.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data with duplicates in column B (index 1)
        string[] sampleData = { "Apple", "Orange", "Apple", "Banana", "Orange", "Grape", "Apple" };
        for (int i = 0; i < sampleData.Length; i++)
        {
            cells[i, 1].PutValue(sampleData[i]); // Row i, Column B
        }

        // Add a conditional formatting rule to highlight duplicate values in column B
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection conditions = worksheet.ConditionalFormattings[cfIndex];

        // Define the range that the conditional formatting will apply to (column B rows)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = sampleData.Length - 1,
            StartColumn = 1,
            EndColumn = 1
        };
        conditions.AddArea(area);

        // Add a DuplicateValues condition
        int conditionIndex = conditions.AddCondition(FormatConditionType.DuplicateValues);
        FormatCondition duplicateCondition = conditions[conditionIndex];

        // Create a style to highlight duplicates (light red background)
        Style highlightStyle = workbook.CreateStyle();
        highlightStyle.ForegroundColor = Color.LightSalmon;
        highlightStyle.Pattern = BackgroundType.Solid;
        duplicateCondition.Style = highlightStyle;

        // Save the workbook with the conditional formatting applied
        workbook.Save("DuplicateValuesHighlight.xlsx");
    }
}

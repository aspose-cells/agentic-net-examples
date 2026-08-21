// Title: C# – Highlight Duplicate Values in a Named Range with Aspose.Cells Conditional Formatting
// Description: This Aspose.Cells for .NET example creates a workbook, fills column A with sample data, defines a named range, and adds a DuplicateValues conditional‑formatting rule that colors duplicate cells yellow. The workbook is saved as HighlightedDuplicates.xlsx.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | duplicate values | named range | highlight duplicates | Excel automation | CellArea | FormatConditionType.DuplicateValues | sample code
// Common Searches: Aspose.Cells highlight duplicates C# | Conditional formatting duplicate values .NET | Create named range Aspose.Cells | Apply duplicate value rule Excel using Aspose | C# code to color duplicate cells
// Developer Intent: The developer wants to automatically detect and visually mark duplicate entries inside a specific named range of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Detect repeated product SKUs in inventory sheets | Identify duplicate customer emails during data import | Flag repeated transaction IDs in financial reports | Assist data‑cleansing by marking duplicate rows | Provide quick visual audit of survey responses
// AI Prompts: Write C# Aspose.Cells code that creates a named range and applies a DuplicateValues conditional format with a red font and bold style. | Explain how to retrieve the cell addresses of duplicates after the conditional formatting is applied using Aspose.Cells. | Show how to replace the solid yellow fill with a two‑color gradient for duplicate cells in Aspose.Cells. | Give a step‑by‑step guide to export the list of duplicate values to a separate worksheet.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This Aspose.Cells for .NET example creates a workbook, fills column A with sample data, defines a named range, and adds a DuplicateValues conditional‑formatting rule that colors duplicate cells yellow. The workbook is saved as HighlightedDuplicates.xlsx.
class HighlightDuplicates
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data with some duplicate values in column A
            string[] sampleData = { "Apple", "Banana", "Apple", "Cherry", "Banana", "Date" };
            for (int i = 0; i < sampleData.Length; i++)
            {
                cells[i, 0].PutValue(sampleData[i]); // A1, A2, ...
            }

            // Define a named range that covers the populated cells (A1:A6)
            AsposeRange namedRange = cells.CreateRange(0, 0, sampleData.Length, 1);
            namedRange.Name = "MyDataRange";

            // Add a conditional formatting rule to highlight duplicate values within the named range
            int cfIndex = worksheet.ConditionalFormattings.Add(); // create a new conditional formatting collection
            var conditionalFormatting = worksheet.ConditionalFormattings[cfIndex];

            // Set the area of the conditional formatting to the named range
            CellArea area = new CellArea
            {
                StartRow = namedRange.FirstRow,
                EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                StartColumn = namedRange.FirstColumn,
                EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
            };
            conditionalFormatting.AddArea(area);

            // Add a condition of type DuplicateValues
            int conditionIndex = conditionalFormatting.AddCondition(FormatConditionType.DuplicateValues);
            var condition = conditionalFormatting[conditionIndex];

            // Define the style to apply to duplicate cells (yellow background)
            Style duplicateStyle = workbook.CreateStyle();
            duplicateStyle.ForegroundColor = Color.Yellow;
            duplicateStyle.Pattern = BackgroundType.Solid;
            condition.Style = duplicateStyle;

            // Save the workbook with highlighted duplicates
            string outputPath = "HighlightedDuplicates.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

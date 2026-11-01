// Title: Apply smart markers and conditional formatting to merged cells with threshold‑based highlighting using Aspose.Cells for .NET (C#)
// Common Searches: how to add a smart marker to a merged cell and apply conditional formatting with Aspose.Cells C# | Aspose.Cells C# conditional formatting based on another cell's value for a merged header | set red fill for cells greater than 100 using Aspose.Cells .NET | highlight merged header yellow when B3 exceeds 150 with Aspose.Cells C# | combine smart markers with expression‑based conditional formatting in Aspose.Cells
// Tags: smart marker insertion merged cell Aspose.Cells | conditional formatting expression based on another cell Aspose.Cells | red fill rule for values greater than 100 Aspose.Cells | yellow fill rule for merged header when B3>150 Aspose.Cells | C# Aspose.Cells example creating output.xlsx

using System;
using System.Drawing;
using Aspose.Cells;

// The code creates a workbook, merges cells B2:D2, inserts a smart marker {Total} into the merged range, populates column B with incremental numbers, adds a red fill conditional format for cells with values >100, adds a yellow fill conditional format for the merged header when B3 >150, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells B2:D2 (zero‑based indices: row 1, column 1 to row 1, column 3)
            sheet.Cells.Merge(1, 1, 1, 3);

            // Insert a variable marker into the merged cell
            Cell mergedCell = sheet.Cells[1, 1];
            mergedCell.PutValue("Total: {Total}");

            // Populate sample data in column B (rows 3‑12) for conditional formatting demo
            for (int i = 2; i < 12; i++) // rows 2‑11 (zero‑based)
            {
                sheet.Cells[i, 1].PutValue((i - 1) * 20); // values: 20,40,...200
            }

            // -------------------------------------------------
            // Conditional Formatting #1: Highlight cells > 100
            // -------------------------------------------------
            // Add a new conditional formatting collection
            int cfIndex1 = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection1 = sheet.ConditionalFormattings[cfIndex1];

            // Define the target range B3:B12 (rows 2‑11, column 1)
            CellArea range1 = new CellArea { StartRow = 2, StartColumn = 1, EndRow = 11, EndColumn = 1 };
            cfCollection1.AddArea(range1);

            // Add a condition: cell value greater than 100
            int conditionIdx1 = cfCollection1.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "100",
                null);
            FormatCondition condition1 = cfCollection1[conditionIdx1];

            // Set the style for the condition (solid red fill)
            Style style1 = workbook.CreateStyle();
            style1.ForegroundColor = Color.Red;
            style1.Pattern = BackgroundType.Solid;
            condition1.Style = style1;

            // -------------------------------------------------
            // Conditional Formatting #2: Highlight merged cell if B3 > 150
            // -------------------------------------------------
            int cfIndex2 = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection2 = sheet.ConditionalFormattings[cfIndex2];

            // Target range is the merged cell B2:D2
            CellArea mergedArea = new CellArea { StartRow = 1, StartColumn = 1, EndRow = 1, EndColumn = 3 };
            cfCollection2.AddArea(mergedArea);

            // Add an expression condition referencing B3 (cell B3 is row 2, column 1)
            int conditionIdx2 = cfCollection2.AddCondition(
                FormatConditionType.Expression,
                OperatorType.None,
                "B3>150",
                null);
            FormatCondition condition2 = cfCollection2[conditionIdx2];

            // Set the style for the condition (solid yellow fill)
            Style style2 = workbook.CreateStyle();
            style2.ForegroundColor = Color.Yellow;
            style2.Pattern = BackgroundType.Solid;
            condition2.Style = style2;

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

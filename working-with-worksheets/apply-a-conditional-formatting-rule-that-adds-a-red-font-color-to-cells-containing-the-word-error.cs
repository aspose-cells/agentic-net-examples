// Title: Add red font conditional formatting for cells that contain the word “Error” using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to format cells with red font when they contain the word Error | apply text‑based conditional formatting in an Excel workbook with Aspose.Cells .NET | Aspose.Cells example for highlighting cells that contain specific text | set font color based on cell text using Aspose.Cells in C# | conditional formatting range A1:C10 for word Error Aspose.Cells tutorial
// Tags: text content conditional formatting Aspose.Cells | red font style conditional formatting Aspose.Cells | range specific conditional formatting Aspose.Cells | Excel cell text highlight Aspose.Cells | C# conditional formatting for specific word

using Aspose.Cells;
using System;
using System.Drawing;

// The code creates a new workbook, defines a conditional formatting rule for the range A1:C10 that flags cells containing the word “Error”, sets their font color to red, and saves the file as ConditionalFormatting_Output.xlsx.
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

            // Define the range for conditional formatting (A1:C10)
            CellArea area = CellArea.CreateCellArea("A1", "C10");

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Add a condition that checks if the cell contains the word "Error"
            // OperatorType.None is used because ContainsText does not require an operator.
            int conditionIndex = cfCollection.AddCondition(
                FormatConditionType.ContainsText,
                OperatorType.None,
                "Error",
                string.Empty);

            // Retrieve the created condition
            FormatCondition condition = cfCollection[conditionIndex];

            // Set the font color of the cells that meet the condition to red
            condition.Style.Font.Color = Color.Red;

            // Apply the defined range to the conditional formatting collection
            cfCollection.AddArea(area);

            // Save the workbook
            workbook.Save("ConditionalFormatting_Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

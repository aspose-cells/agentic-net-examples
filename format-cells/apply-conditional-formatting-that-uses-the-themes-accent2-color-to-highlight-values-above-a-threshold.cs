// Title: Aspose.Cells .NET – Conditional Formatting with Accent2 Theme Color for Values Over a Threshold
// Description: Shows how to create a workbook, populate column A with sample numbers, add a conditional formatting rule for A1:A10 that highlights cells greater than 50 using the Accent2 theme color (RGB 0,176,240) and bold text, and save the result as an .xlsx file.
// Keywords: Aspose.Cells | .NET | C# | conditional formatting | Accent2 theme color | highlight cells above threshold | Excel solid fill | cell style formatting | threshold based coloring
// Common Searches: Aspose.Cells conditional formatting Accent2 C# | how to highlight cells greater than a value using Aspose.Cells | apply theme accent color to conditional format .NET | set solid background for conditional rule in Aspose.Cells | C# example conditional formatting threshold Excel
// Developer Intent: Create a conditional formatting rule that colors cells with values exceeding a given threshold using the workbook’s Accent2 theme color.
// Use Cases: Flag sales figures that surpass a target amount in a financial dashboard. | Mark temperature readings above safety limits in an engineering log. | Highlight overdue task counts that exceed a defined threshold in a project tracker.
// AI Prompts: Write C# code with Aspose.Cells to apply Accent2 theme color conditional formatting for cells greater than a specified number. | Modify the example to use a different theme accent (e.g., Accent3) and a custom threshold value. | Extend the solution to apply the same Accent2 conditional formatting to multiple columns or a whole table.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Retained for potential future use

// Shows how to create a workbook, populate column A with sample numbers, add a conditional formatting rule for A1:A10 that highlights cells greater than 50 using the Accent2 theme color (RGB 0,176,240) and bold text, and save the result as an .xlsx file.
class ConditionalFormattingAccent2
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (rows 1-10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // Values: 0,10,20,...,90
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the formatting will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add a CellValue condition: highlight cells greater than the threshold (e.g., 50)
            int conditionIndex = fcc.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",   // Formula1 – the threshold value
                null); // Formula2 – not used for GreaterThan

            // Retrieve the created condition
            FormatCondition condition = fcc[conditionIndex];

            // Apply a solid fill using a color that matches the typical Accent2 theme color
            condition.Style.BackgroundColor = System.Drawing.Color.FromArgb(0, 176, 240);
            condition.Style.Pattern = BackgroundType.Solid;

            // Optionally, make the text bold for better visibility
            condition.Style.Font.IsBold = true;

            // Define output file name
            string outputPath = "ConditionalFormatting_Accent2.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

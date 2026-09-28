// Title: Add a greater‑than conditional formatting rule with the theme Accent2 solid fill to a range in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a conditional formatting rule using a solid fill that matches the workbook's Accent2 theme when cell values exceed a threshold. | Write a C# snippet that applies a solid Accent2 background to cells in range A2:A11 via Aspose.Cells conditional formatting for values greater than 60.
// Common Searches: Aspose.Cells C# apply Excel Accent2 theme color in conditional formatting | C# Aspose.Cells conditional formatting greater than threshold with solid fill | set conditional formatting fill to Excel theme Accent2 using Aspose.Cells .NET
// Tags: aspocells conditional formatting accent2 fill | c# conditional formatting greater than rule | excel theme color conditional format aspocells | apply conditional formatting to range a2:a11 c#

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, fills column A with sample numbers, defines the range A2:A11, adds a conditional formatting rule that highlights cells with values above 60 using a solid fill approximating the theme's Accent2 color, and saves the file as ConditionalFormatting_Accent2.xlsx.
class ConditionalFormattingExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (rows 2 to 11)
            for (int i = 0; i < 10; i++)
            {
                // Example values ranging from 30 to 80
                sheet.Cells[i + 1, 0].PutValue(30 + i * 5);
            }

            // Define the range to which the conditional formatting will be applied (A2:A11)
            CellArea area = new CellArea
            {
                StartRow = 1,   // zero‑based index (row 2)
                EndRow = 10,    // row 11
                StartColumn = 0,
                EndColumn = 0
            };

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();

            // Use dynamic to avoid compile‑time dependency on ConditionalFormatting type
            dynamic cf = sheet.ConditionalFormattings[cfIndex];

            // Associate the defined range with this conditional formatting
            cf.AddArea(area);

            // Add a condition: highlight cells with values greater than the threshold (e.g., 60)
            double threshold = 60.0;
            int conditionIndex = cf.AddCondition(
                FormatConditionType.CellValue,   // condition type
                OperatorType.GreaterThan,        // operator
                threshold.ToString(),            // first formula (threshold)
                null);                           // second formula (not used)

            // Create a style that uses a solid fill (simulating theme Accent2 color)
            Style accentStyle = workbook.CreateStyle();
            accentStyle.ForegroundColor = Color.FromArgb(0, 176, 240); // Approximate Accent2
            accentStyle.Pattern = BackgroundType.Solid;               // solid fill

            // Assign the style to the condition
            FormatCondition condition = cf[conditionIndex];
            condition.Style = accentStyle;

            // Save the workbook
            workbook.Save("ConditionalFormatting_Accent2.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

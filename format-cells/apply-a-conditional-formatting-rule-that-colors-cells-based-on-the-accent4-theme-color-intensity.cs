// Title: C# – Aspose.Cells: Apply Accent4 Theme Color Scale Conditional Formatting to a Cell Range
// Description: Creates a workbook, fills cells A1:A10 with values 1‑10, adds a three‑point ColorScale conditional formatting rule, and assigns darkened, normal, and lightened Accent4 theme colors using ThemeColor tints before saving the file as an .xlsx document.
// Keywords: Aspose.Cells | C# conditional formatting | Accent4 theme color | color scale | theme tint | FormatConditionType.ColorScale | ThemeColorType.Accent4 | Excel gradient formatting | Aspose.Cells example | conditional formatting API
// Common Searches: Aspose.Cells color scale Accent4 | C# conditional formatting theme tint | set Accent4 darken lighten Aspose.Cells | three‑point color scale with theme colors .NET | apply percentile mid color using Aspose.Cells | ThemeColor conditional formatting C#
// Developer Intent: Generate a three‑point color‑scale conditional formatting rule that uses the Accent4 theme color with dark, normal, and light tints for cells A1:A10.
// Use Cases: Visualize sales or KPI data with a gradient from dark to light Accent4, highlighting low‑to‑high values in reports. | Maintain brand consistency in Excel dashboards by applying the workbook’s theme colors to numeric columns. | Enhance financial models with a percentile‑based mid color that emphasizes median values while using theme‑tinted gradients.
// AI Prompts: Show how to modify the code to use the Accent2 theme color while keeping the same darken/lighten tints. | Generate C# code for a two‑color conditional formatting rule that applies a single theme color with custom dark and light tints. | Explain how to read back the ThemeColor values from an existing conditional formatting rule after the workbook is saved.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ConditionalFormattingAccent4Demo
{
    // Creates a workbook, fills cells A1:A10 with values 1‑10, adds a three‑point ColorScale conditional formatting rule, and assigns darkened, normal, and lightened Accent4 theme colors using ThemeColor tints before saving the file as an .xlsx document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // Values 1..10
            }

            // Add an empty conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add a ColorScale conditional formatting rule
            fcs.AddCondition(FormatConditionType.ColorScale);
            FormatCondition fc = fcs[0]; // The newly added condition

            // Configure the three color points (Min, Mid, Max) for the color scale
            // Minimum point – darkened Accent4
            fc.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;
            CellsColor minColorCell = workbook.CreateCellsColor();
            minColorCell.ThemeColor = new ThemeColor(ThemeColorType.Accent4, -0.5); // 50% darken
            fc.ColorScale.MinColor = minColorCell.Color;

            // Mid point – normal Accent4
            fc.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile;
            fc.ColorScale.MidCfvo.Value = 50; // 50th percentile
            CellsColor midColorCell = workbook.CreateCellsColor();
            midColorCell.ThemeColor = new ThemeColor(ThemeColorType.Accent4, 0.0); // no tint
            fc.ColorScale.MidColor = midColorCell.Color;

            // Maximum point – lightened Accent4
            fc.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;
            CellsColor maxColorCell = workbook.CreateCellsColor();
            maxColorCell.ThemeColor = new ThemeColor(ThemeColorType.Accent4, 0.5); // 50% lighten
            fc.ColorScale.MaxColor = maxColorCell.Color;

            // Save the workbook
            workbook.Save("ConditionalFormattingAccent4.xlsx");
        }
    }
}

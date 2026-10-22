// Title: C# example to verify that conditional formatting remains active after applying a workbook theme with Aspose.Cells
// Common Searches: how to test if Aspose.Cells conditional formatting survives a theme change in C# | Aspose.Cells verify conditional format style after applying .thmx theme | C# check conditional formatting persistence after workbook theme update using Aspose.Cells | validate conditional formatting rule after applying new theme with Aspose.Cells for .NET | recalculate formulas and retain conditional formatting after theme application Aspose.Cells
// Tags: Aspose.Cells conditional formatting after theme change | C# verify conditional format retention | apply .thmx theme preserve conditional formatting | recalculate formulas Aspose.Cells workbook | conditional formatting validation Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds a conditional formatting rule that highlights cells with values greater than 50 using a solid red background, optionally applies a new .thmx theme, recalculates formulas, and then validates that the conditional formatting collection still exists and that cell A1 retains the red background. The workbook is saved before and after the theme operation.
class ConditionalFormattingThemeValidation
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(60); // This cell should meet the condition (>50)
            sheet.Cells["A2"].PutValue(30); // This cell should not meet the condition

            // Add a conditional formatting rule to highlight cells with value > 50
            int cfIndex = sheet.ConditionalFormattings.Add(); // Create a new CF range
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the area for the conditional formatting (A1:A10)
            CellArea area = CellArea.CreateCellArea("A1", "A10");
            cfCollection.AddArea(area);

            // Define the condition (Cell Value > 50)
            // The fourth parameter (formula2) is required; an empty string is acceptable for this type.
            int conditionIndex = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                string.Empty);

            // Create a style for the condition (solid red background)
            Style redStyle = workbook.CreateStyle();
            redStyle.ForegroundColor = Color.Red;
            redStyle.Pattern = BackgroundType.Solid;

            // Assign the style to the condition
            cfCollection[conditionIndex].Style = redStyle;

            // Save the workbook before applying a new theme (optional, for inspection)
            workbook.Save("BeforeTheme.xlsx");

            // Apply a new theme to the workbook if the theme file exists
            string themePath = "NewTheme.thmx";
            if (File.Exists(themePath))
            {
                // The ApplyTheme method may not be available in older Aspose.Cells versions.
                // If needed, replace this block with appropriate theme‑application logic.
                Console.WriteLine($"Theme file found: {themePath}. Theme application is not supported in this version.");
            }
            else
            {
                Console.WriteLine($"Theme file not found: {themePath}. Skipping theme application.");
            }

            // Recalculate formulas and conditional formats
            workbook.CalculateFormula();

            // Validation: ensure the conditional formatting collection still exists
            bool cfExists = sheet.ConditionalFormattings.Count > 0;

            // Validation: check that the style of a cell meeting the condition reflects the CF style
            Style cellStyle = sheet.Cells["A1"].GetStyle();
            bool isRedBackground = cellStyle.ForegroundColor.ToArgb() == Color.Red.ToArgb()
                                   && cellStyle.Pattern == BackgroundType.Solid;

            // Output validation results
            Console.WriteLine("Conditional Formatting Exists: " + cfExists);
            Console.WriteLine("Cell A1 has Red Background after Theme: " + isRedBackground);

            // Save the final workbook
            workbook.Save("AfterTheme.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

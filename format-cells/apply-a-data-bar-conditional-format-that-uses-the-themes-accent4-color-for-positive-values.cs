// Title: Create a data bar conditional format with the theme Accent4 color for positive values using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add data bar conditional formatting with theme accent color | How to set DataBar.Color to Accent4 using GetThemeColor in Aspose.Cells | Apply data bar conditional format to range A1:A10 with Aspose.Cells .NET | Use theme colors for conditional formatting data bars in C# Aspose.Cells | Configure DataBar.ShowValue and theme color in Aspose.Cells workbook
// Tags: data bar conditional formatting Aspose.Cells C# | theme accent color GetThemeColor Aspose.Cells | apply conditional formatting to cell range Aspose.Cells | set DataBar.Color using theme Accent4 | save workbook with data bar formatting Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, defines the range A1:A10, adds a DataBar conditional formatting rule, enables ShowValue, sets the bar color to the workbook's theme Accent4 via GetThemeColor, applies the rule to the range, and saves the file as DataBar_Accent4.xlsx.
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

            // Define the range to which the data bar will be applied (e.g., A1:A10)
            CellArea range = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a new conditional formatting rule to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();

            // Retrieve the conditional formatting object (using var to avoid type issues)
            var cf = sheet.ConditionalFormattings[cfIndex];

            // Add a DataBar condition to the rule
            cf.AddCondition(FormatConditionType.DataBar);
            FormatCondition dataBarCondition = cf[0];

            // Configure the DataBar
            DataBar dataBar = dataBarCondition.DataBar;
            dataBar.ShowValue = true; // Show the cell value next to the bar
            dataBar.Color = workbook.GetThemeColor(ThemeColorType.Accent4); // Use theme Accent4 for positive values

            // Apply the conditional formatting to the defined range
            cf.AddArea(range);

            // Save the workbook
            string outputPath = "DataBar_Accent4.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

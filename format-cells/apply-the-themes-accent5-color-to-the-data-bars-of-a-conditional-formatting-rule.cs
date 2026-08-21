// Title: Aspose.Cells for .NET – Apply Theme Accent5 Color to a Data Bar Conditional Formatting Rule
// Description: Creates a workbook, fills cells A1:A5 with numbers, adds a Data Bar conditional format, sets automatic min/max, shows values, and assigns the workbook's Accent5 theme color to the data bar before saving as XLSX.
// Keywords: Aspose.Cells C# data bar color | theme accent color conditional formatting | Workbook.GetThemeColor | Accent5 data bar Aspose | conditional formatting Excel .NET
// Common Searches: set data bar color using workbook theme in Aspose.Cells | apply Accent5 theme color to conditional formatting data bar C# | Aspose.Cells GetThemeColor example | change Excel data bar color programmatically .NET
// Developer Intent: Assign the workbook’s Accent5 theme color to a Data Bar conditional formatting rule using Aspose.Cells for .NET.
// Use Cases: Generate reports where data bars match the document’s theme for consistent branding. | Programmatically style multiple worksheets so all data bars share the same accent color. | Build dashboards that automatically adapt data‑bar colors when the workbook theme changes.
// AI Prompts: Show how to retrieve a theme color and apply it to a DataBar conditional format in Aspose.Cells C#. | Provide code to change an existing DataBar rule to use the workbook’s Accent2 color. | Explain how to apply different theme accent colors to several conditional formatting rules in one workbook.

using System;
using Aspose.Cells;
using System.Drawing;

// Creates a workbook, fills cells A1:A5 with numbers, adds a Data Bar conditional format, sets automatic min/max, shows values, and assigns the workbook's Accent5 theme color to the data bar before saving as XLSX.
class ApplyAccent5ToDataBar
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (rows 1-5)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(30);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(40);
            sheet.Cells["A5"].PutValue(25);

            // Define the cell area that will receive the conditional formatting (A1:A5)
            CellArea area = new CellArea { StartRow = 0, EndRow = 4, StartColumn = 0, EndColumn = 0 };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
            cfCollection.AddArea(area);

            // Add a DataBar condition to the collection
            int conditionIdx = cfCollection.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = cfCollection[conditionIdx];

            // Configure the DataBar (automatic min/max and show cell values)
            condition.DataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
            condition.DataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
            condition.DataBar.ShowValue = true;

            // Apply the workbook theme's Accent5 color to the data bar
            condition.DataBar.Color = workbook.GetThemeColor(ThemeColorType.Accent5);

            // Save the workbook to a file
            string outputPath = "Accent5DataBar.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

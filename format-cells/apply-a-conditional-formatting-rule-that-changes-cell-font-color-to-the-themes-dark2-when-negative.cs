// Title: C# – Aspose.Cells: Conditional Formatting to Apply Dark2 Theme Font Color for Negative Values
// Description: Creates a workbook, fills cells A1:A4 with mixed numbers, adds a conditional formatting rule for the range, triggers when the cell value is less than zero, retrieves the workbook's Dark2 (Accent2) theme color, applies it to the font style, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# conditional formatting | negative values font color | Dark2 theme color | Accent2 Excel theme | .NET Excel styling | GetThemeColor | FormatCondition | Excel conditional formatting API | programmatic Excel formatting
// Common Searches: Aspose.Cells set font color for negative numbers | C# apply Dark2 theme color with conditional formatting | How to use GetThemeColor in Aspose.Cells | Conditional formatting rule less than zero Aspose.Cells | Change font color based on cell value in .NET Excel
// Developer Intent: Add a conditional formatting rule that colors the font of cells with negative values using the workbook’s Dark2 (Accent2) theme color.
// Use Cases: Highlight financial losses in reports with a consistent Dark2 accent. | Visually separate negative KPI metrics from positive ones in dashboards. | Apply uniform theme‑based negative‑value styling across multiple worksheets in a single workbook.
// AI Prompts: Generate C# code with Aspose.Cells that formats negative numbers using the Dark2 theme font color. | Show how to retrieve the Dark2 (Accent2) theme color from a workbook and apply it in a conditional formatting rule. | Provide an Aspose.Cells example that creates a workbook, adds a less‑than‑zero rule, sets the font color to Dark2, and saves the file, ensuring the output directory exists.

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// Creates a workbook, fills cells A1:A4 with mixed numbers, adds a conditional formatting rule for the range, triggers when the cell value is less than zero, retrieves the workbook's Dark2 (Accent2) theme color, applies it to the font style, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with positive and negative values
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(-5);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(-15);

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the range for conditional formatting (A1:A4)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 3,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add a condition for cells with values less than 0 (negative values)
            int condIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "0", string.Empty);
            FormatCondition fc = fcs[condIdx];

            // Set the font color for negative values to a theme color (using Accent1 as an example)
            Color themeColor = workbook.GetThemeColor(ThemeColorType.Accent1);
            fc.Style.Font.Color = themeColor;

            // Define output file path
            string outputPath = "NegativeFontColor.xlsx";

            // Ensure the directory exists (if a directory part is present)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

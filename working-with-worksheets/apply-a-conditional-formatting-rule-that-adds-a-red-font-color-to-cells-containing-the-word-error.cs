// Title: How to apply red font conditional formatting for cells containing the word "Error" using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# conditional formatting to color text red when a cell contains 'Error' | Set FormatConditionType.ContainsText with a red font style in a .NET workbook | Apply conditional formatting to a range of cells for error messages using Aspose.Cells | Create red font style for cells that include the word Error in an Excel file with Aspose.Cells | C# code to highlight error text with red font via conditional formatting in an XLSX workbook
// Tags: Aspose.Cells conditional formatting contains text | red font style FormatConditionType.ContainsText | C# apply conditional formatting to Excel range | highlight error cells Aspose.Cells | CreateStyle font color red Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, defines a cell area covering the first 100 rows and 10 columns, adds a conditional formatting collection, and configures a FormatCondition of type ContainsText with the text "Error". A style with a red font color is applied when the condition matches, and the workbook is saved as ConditionalFormatting.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range where the conditional formatting will be applied
            // Here we apply it to the first 100 rows and first 10 columns (adjust as needed)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 99,
                EndColumn = 9
            };

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Associate the defined range with this conditional formatting
            fcc.AddArea(area);

            // Create a condition that checks if the cell contains the word "Error"
            FormatCondition condition = fcc[0];
            condition.Type = FormatConditionType.ContainsText; // Correct enum
            condition.Text = "Error";

            // Define the style to apply when the condition is met (red font color)
            Style redFontStyle = workbook.CreateStyle();
            redFontStyle.Font.Color = Color.Red;
            condition.Style = redFontStyle;

            // Save the workbook with the conditional formatting applied
            string outputPath = "ConditionalFormatting.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

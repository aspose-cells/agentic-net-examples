// Title: Export an Excel workbook with a three‑color ColorScale conditional format to HTML using Aspose.Cells for .NET, embedding the CSS in a <style> tag
// AI Prompts: Generate C# code that creates a workbook, applies a three‑color ColorScale (light green, yellow, red) to the range A1:A10, and saves it as HTML with the conditional formatting rendered inside a <style> block via Aspose.Cells. | Modify the example to use a custom ColorScale of LightBlue, Orange, and DarkRed, then verify that the produced HTML file contains the corresponding CSS rules in the <style> section. | Add a data‑bar conditional formatting rule to the same range while preserving the existing ColorScale, and export the workbook to HTML ensuring both formatting rules appear in the generated style block.
// Common Searches: aspnet export excel with color scale conditional formatting to html using aspose.cells | how to embed conditional formatting CSS when saving workbook as html with aspose.cells | c# generate three color scale conditional format and save as html file | asp.net save workbook as html and include conditional formatting styles in style tag
// Tags: Aspose.Cells HTMLSaveOptions conditional formatting | C# export Excel workbook to HTML | conditional formatting gradient Aspose.Cells | embed conditional formatting CSS in HTML output | add data bar rule to HTML export

using System;
using System.Drawing;
using Aspose.Cells;

// // This program creates a new workbook, fills cells A1:A10 with values 1‑10, applies a three‑color ColorScale (light green → yellow → red) conditional formatting, and saves the workbook as an HTML file where Aspose.Cells emits the conditional‑formatting rules as CSS inside a <style> block.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A10).
            for (int row = 0; row < 10; row++)
            {
                sheet.Cells[row, 0].PutValue(row + 1);
            }

            // Add a ColorScale conditional formatting to the range A1:A10.
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 });

            // Add a three‑color scale condition.
            int conditionIndex = cf.AddCondition(FormatConditionType.ColorScale);
            var condition = cf[conditionIndex];

            // Set the colors for the scale. The default types are Min, Percentile (50), and Max.
            condition.ColorScale.MinColor = Color.LightGreen;
            condition.ColorScale.MidColor = Color.Yellow;
            condition.ColorScale.MaxColor = Color.Red;

            // Save the workbook as an HTML file.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            workbook.Save("ConditionalColorScale.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

// Title: Export a workbook with 3‑color ColorScale conditional formatting to a single HTML file and embed the CSS in a <style> block (Aspose.Cells for .NET)
// Description: This C# sample builds a 10×10 worksheet, applies a three‑color ColorScale (red‑yellow‑green) via conditional formatting, sets min, percentile, and max thresholds, and saves the workbook as one HTML document with the gradient defined in an inline <style> element using HtmlSaveOptions.
// Keywords: Aspose.Cells | .NET | C# | HTML export | conditional formatting | ColorScale | 3‑color gradient | embedded CSS | single file HTML | Excel to HTML | CssStyles property
// Common Searches: Aspose.Cells export Excel to HTML with ColorScale | How to embed CSS when saving workbook as HTML in .NET | Create 3‑color conditional formatting in Aspose.Cells | Save workbook as single HTML file Aspose | Add custom style block to HTML output Aspose.Cells
// Developer Intent: Produce an HTML representation of an Excel workbook that retains a three‑color ColorScale rule and contains the required CSS directly inside the page.
// Use Cases: Web‑based reporting where cell values are color‑graded without external style sheets | Email‑ready HTML snapshot of an Excel sheet preserving visual cues | Dashboard component that loads a pre‑styled HTML file generated from Excel data
// AI Prompts: Generate code to switch the ColorScale to a 2‑color (red‑green) scheme and adjust the inline CSS. | Show how to extract the generated CSS class name and add font‑color or border styling. | Demonstrate exporting multiple worksheets, each with distinct conditional formats, while consolidating all CSS into one <style> block.

using System.Drawing;
using Aspose.Cells;

// This C# sample builds a 10×10 worksheet, applies a three‑color ColorScale (red‑yellow‑green) via conditional formatting, sets min, percentile, and max thresholds, and saves the workbook as one HTML document with the gradient defined in an inline <style> element using HtmlSaveOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with sample numeric data
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                worksheet.Cells[row, col].PutValue(row * col);
            }
        }

        // Add a ColorScale conditional formatting rule (3‑color: Red → Yellow → Green)
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

        // Define the range to which the conditional formatting applies
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 9
        };
        fcs.AddArea(area);

        // Create the ColorScale condition
        int conditionIndex = fcs.AddCondition(FormatConditionType.ColorScale);
        FormatCondition fc = fcs[conditionIndex];

        // Configure the ColorScale (min = Red, mid = Yellow, max = Green)
        fc.ColorScale.Is3ColorScale = true;
        fc.ColorScale.MinColor = Color.Red;
        fc.ColorScale.MidColor = Color.Yellow;
        fc.ColorScale.MaxColor = Color.Green;

        // Set the value objects for the scale
        fc.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;
        fc.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile;
        fc.ColorScale.MidCfvo.Value = 50; // 50th percentile
        fc.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;

        // Prepare HTML save options to embed CSS in a <style> block
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.SaveAsSingleFile = true;               // All content in one HTML file
        htmlOptions.CssStyles = @"
            /* Custom CSS for the ColorScale conditional formatting */
            .cs0 { background: linear-gradient(to right, red, yellow, green); }
        ";

        // Save the workbook as HTML with the embedded style definitions
        workbook.Save("ColorScaleWithStyle.html", htmlOptions);
    }
}

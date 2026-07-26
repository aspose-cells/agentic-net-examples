// Title: Export Excel with DataBar Conditional Formatting to HTML and External CSS using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply an orange DataBar conditional format with a solid black border to cells A1:A10, and save the sheet as HTML while generating a separate CSS stylesheet via HtmlSaveOptions.ExportWorksheetCSSSeparately.
// Keywords: Aspose.Cells | C# | HTML export | external CSS | DataBar conditional formatting | ExportWorksheetCSSSeparately | Excel to HTML | worksheet CSS file | conditional formatting to HTML | Aspose.Cells example
// Common Searches: Aspose.Cells export DataBar to HTML | C# save Excel as HTML with external stylesheet | HtmlSaveOptions ExportWorksheetCSSSeparately example | How to generate separate CSS file from Aspose.Cells | DataBar conditional formatting HTML output
// Developer Intent: Export a workbook that contains a DataBar conditional format to HTML while writing the styling to an external CSS file.
// Use Cases: Build web reports that reuse a single CSS file for multiple worksheets with DataBar visuals. | Create HTML email templates where conditional formatting is applied via external styles for better client compatibility. | Develop dashboards that separate content (HTML) from presentation (CSS) for caching and faster page loads.
// AI Prompts: Generate C# code that adds a gradient DataBar to a range and saves the worksheet as HTML with an external CSS file using Aspose.Cells. | Explain how to edit the generated CSS file to change the DataBar color after the HTML export. | Show the steps to load an existing workbook, apply a DataBar conditional format to a specific range, and export it to HTML with a separate stylesheet.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a workbook, apply an orange DataBar conditional format with a solid black border to cells A1:A10, and save the sheet as HTML while generating a separate CSS stylesheet via HtmlSaveOptions.ExportWorksheetCSSSeparately.
class ExportDataBarWithSeparateCss
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate column A with sample numeric data
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i * 10); // A1..A10
        }

        // Add an empty conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

        // Define the cell area (A1:A10) for the DataBar conditional formatting
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        cfCollection.AddArea(area);

        // Add a DataBar condition to the collection
        int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
        FormatCondition condition = cfCollection[conditionIndex];

        // Configure the DataBar properties
        DataBar dataBar = condition.DataBar;
        dataBar.Color = Color.Orange;                         // Bar color
        dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum automatically
        dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum automatically
        dataBar.ShowValue = true;                             // Show cell values
        dataBar.BarBorder.Type = DataBarBorderType.Solid;     // Solid border
        dataBar.BarBorder.Color = Color.Black;                // Border color
        dataBar.BarFillType = DataBarFillType.Solid;          // Solid fill

        // Set HTML save options to export worksheet CSS separately
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true;      // Generates a separate CSS file

        // Save the workbook as HTML; the CSS file will be created alongside the HTML file
        string htmlPath = "DataBarOutput.html";
        workbook.Save(htmlPath, saveOptions);
    }
}

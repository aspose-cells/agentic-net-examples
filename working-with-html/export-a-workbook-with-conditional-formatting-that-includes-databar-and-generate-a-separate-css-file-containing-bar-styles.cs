// Title: Export Excel with DataBar Conditional Formatting to HTML and Separate CSS using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills column A with numeric values, applies an orange DataBar with a black solid border to A1:A10, and saves the sheet as HTML. HtmlSaveOptions.ExportWorksheetCSSSeparately generates an external CSS file that contains the DataBar styles.
// Keywords: Aspose.Cells HTML export | DataBar conditional formatting C# | ExportWorksheetCSSSeparately | separate CSS file Aspose.Cells | Excel to HTML with DataBar
// Common Searches: Aspose.Cells export DataBar to HTML | C# generate external CSS for Excel conditional formatting | HtmlSaveOptions ExportWorksheetCSSSeparately example | How to save Excel as HTML with separate stylesheet | DataBar style CSS Aspose.Cells .NET
// Developer Intent: Generate HTML from an Excel workbook that includes a DataBar conditional format and output the visual styles into a dedicated CSS file.
// Use Cases: Build web‑based reports where DataBars are rendered in HTML and styling is maintained in an external stylesheet for easy updates. | Create reusable CSS assets for conditional formatting across multiple exported worksheets. | Automate HTML dashboards that visualize numeric trends with DataBars while keeping design separation between markup and style.
// AI Prompts: Write C# code with Aspose.Cells to add an orange DataBar to a range and export the worksheet to HTML with an external CSS file. | Show how to modify the generated CSS to change the DataBar color after the HTML export. | Explain how to configure HtmlSaveOptions to include only the styles actually used in the exported HTML.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, fills column A with numeric values, applies an orange DataBar with a black solid border to A1:A10, and saves the sheet as HTML. HtmlSaveOptions.ExportWorksheetCSSSeparately generates an external CSS file that contains the DataBar styles.
class ExportDataBarWithSeparateCss
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample numeric data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i * 10);
        }

        // Add a DataBar conditional formatting rule to the range A1:A10
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

        // Define the cell area for the conditional formatting
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        cfCollection.AddArea(area);

        // Add the DataBar condition
        int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
        FormatCondition condition = cfCollection[conditionIndex];

        // Configure the DataBar appearance
        DataBar dataBar = condition.DataBar;
        dataBar.Color = Color.Orange;                                 // Bar color
        dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum value
        dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum value
        dataBar.ShowValue = true;                                     // Show cell value
        dataBar.BarBorder.Type = DataBarBorderType.Solid;             // Solid border
        dataBar.BarBorder.Color = Color.Black;                        // Border color
        dataBar.BarFillType = DataBarFillType.Solid;                  // Fill type

        // Set HTML save options to export worksheet CSS separately
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true; // Generates a separate .css file
        // Optional: keep all styles (set to false if you want to exclude unused styles)
        saveOptions.ExcludeUnusedStyles = false;

        // Save the workbook as HTML; a corresponding CSS file will be created automatically
        string htmlFilePath = "DataBarExport.html";
        workbook.Save(htmlFilePath, saveOptions);

        Console.WriteLine($"HTML file saved to: {htmlFilePath}");
        Console.WriteLine("A separate CSS file containing the DataBar styles has been generated.");
    }
}

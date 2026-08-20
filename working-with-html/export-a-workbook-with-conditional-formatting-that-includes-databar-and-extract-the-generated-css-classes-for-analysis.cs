// Title: Export DataBar Conditional Formatting to HTML and Retrieve Generated CSS with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply a green DataBar conditional format, save the sheet as HTML using the ExportWorksheetCSSSeparately option, and read the resulting stylesheet.css for analysis. The example is written in C# and uses Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | HTML export | DataBar conditional formatting | ExportWorksheetCSSSeparately | extract CSS | stylesheet.css | Excel to HTML | external stylesheet | conditional formatting to HTML
// Common Searches: Aspose.Cells export DataBar to HTML | How to get CSS file from Aspose.Cells HTML export | ExportWorksheetCSSSeparately example C# | Read stylesheet.css after saving workbook as HTML | DataBar conditional format HTML output Aspose.Cells
// Developer Intent: Generate HTML with a separate CSS file from a workbook that contains a DataBar conditional format and read the CSS content programmatically.
// Use Cases: Create web‑ready reports that show DataBar visual cues while keeping styles in an external stylesheet for easy maintenance. | Programmatically analyze or modify the CSS classes produced by Aspose.Cells to customize DataBar appearance after export. | Batch‑process multiple Excel files with DataBar formatting, export them to HTML, and consolidate their CSS for a unified web UI.
// AI Prompts: Write C# code using Aspose.Cells to add a green DataBar conditional format, export the worksheet to HTML with ExportWorksheetCSSSeparately enabled, and output the contents of the generated stylesheet.css. | Explain the folder structure created by ExportWorksheetCSSSeparately and how to locate the stylesheet.css file after saving a workbook as HTML. | Show how to change the DataBar color, set custom minimum and maximum values, and toggle the value display before exporting to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// Demonstrates how to create a workbook, apply a green DataBar conditional format, save the sheet as HTML using the ExportWorksheetCSSSeparately option, and read the resulting stylesheet.css for analysis. The example is written in C# and uses Aspose.Cells for .NET.
class DataBarHtmlExport
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i + 1);
        }

        // Add an empty conditional formatting collection
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

        // Define the range for the DataBar (A1:A10)
        CellArea area = new CellArea
        {
            StartRow = 0,
            EndRow = 9,
            StartColumn = 0,
            EndColumn = 0
        };
        fcc.AddArea(area);

        // Add a DataBar condition to the collection
        int conditionIndex = fcc.AddCondition(FormatConditionType.DataBar);
        FormatCondition condition = fcc[conditionIndex];

        // Configure the DataBar properties
        DataBar dataBar = condition.DataBar;
        dataBar.Color = Color.Green;                                 // Bar color
        dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum value
        dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum value
        dataBar.ShowValue = true;                                    // Show cell values on the bar

        // Set HTML save options to export worksheet CSS separately
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Define output directory and ensure it exists
        string outputDir = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
        Directory.CreateDirectory(outputDir);

        // Path for the generated HTML file
        string htmlPath = Path.Combine(outputDir, "DataBar.html");

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlPath, saveOptions);

        // When ExportWorksheetCSSSeparately is true, Aspose.Cells creates a subfolder
        // named "<htmlFileName>_files" containing "stylesheet.css"
        string cssPath = Path.Combine(outputDir, "DataBar_files", "stylesheet.css");

        // Read and output the generated CSS classes for analysis
        if (File.Exists(cssPath))
        {
            string cssContent = File.ReadAllText(cssPath);
            Console.WriteLine("=== Extracted CSS Classes ===");
            Console.WriteLine(cssContent);
        }
        else
        {
            Console.WriteLine("CSS file not found at expected location: " + cssPath);
        }
    }
}

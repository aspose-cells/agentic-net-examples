using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Save the HTML as a single file so that CssStyles are embedded
        saveOptions.SaveAsSingleFile = true;

        // Optional: set a CSS class prefix for cells (helps target specific cells if needed)
        saveOptions.CellCssPrefix = "custom-cell-";

        // Define custom CSS that overrides the default cell background colors
        // The !important flag ensures our styles take precedence over generated ones
        saveOptions.CssStyles = @"
            td { background-color: #e0f7fa !important; }
            .custom-cell- { background-color: #ffeb3b !important; }
        ";

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
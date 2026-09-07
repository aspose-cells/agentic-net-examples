// Title: Enable CSS custom properties and assign a custom TableCssId when exporting a workbook to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells HtmlSaveOptions to export a workbook to HTML with EnableCssCustomProperties enabled and a specific TableCssId applied. | Demonstrate how to configure Aspose.Cells HTML export to leverage CSS variables for styling and set a custom CSS ID on the generated table.
// Common Searches: asp.net aspose.cells html export enable css custom properties | how to set tablecssid in htmlsaveoptions c# | aspose.cells export workbook to html with css variables and custom table id | c# save workbook as html using aspose.cells custom css id | using css custom properties in aspose.cells html output
// Tags: Aspose.Cells HtmlSaveOptions CSS custom properties | Aspose.Cells HtmlSaveOptions TableCssId | C# Aspose.Cells HTML export styling | Aspose.Cells CSS variables for HTML tables

using Aspose.Cells;

// Creates a workbook, fills it with sample data, configures HtmlSaveOptions to enable CSS custom properties and assign a custom TableCssId, then saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Populate the workbook with sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1.2);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(0.8);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Enable CSS custom properties for flexible styling
        saveOptions.EnableCssCustomProperties = true;
        // Assign a custom CSS ID to the generated HTML table
        saveOptions.TableCssId = "myCustomTable";

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}

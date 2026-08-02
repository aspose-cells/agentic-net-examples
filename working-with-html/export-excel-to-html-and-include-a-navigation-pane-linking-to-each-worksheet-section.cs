// Title: Export Excel to a single HTML file with a worksheet navigation pane using Aspose.Cells for .NET
// Description: Shows how to build a workbook with several worksheets, fill them with data, and save the whole workbook as one HTML document that contains a navigation pane (sheet tabs) by setting HtmlSaveOptions: ExportActiveWorksheetOnly = false, SaveAsSingleFile = true, ShowAllSheets = true.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | HtmlSaveOptions | ExportActiveWorksheetOnly | SaveAsSingleFile | ShowAllSheets | navigation pane | sheet tabs | single HTML file | workbook export
// Common Searches: Aspose.Cells export workbook to single HTML with sheet navigation | HtmlSaveOptions ShowAllSheets example C# | Save Excel as HTML with tabs using Aspose.Cells | Generate HTML report with worksheet navigation pane .NET | How to create a navigation pane for Excel sheets in HTML export
// Developer Intent: Create one HTML report from an Excel workbook that lets users jump between worksheets via a clickable navigation pane.
// Use Cases: Deliver a web‑ready report that consolidates multiple worksheets into a single page with tabbed navigation. | Provide an interactive HTML view of a workbook for users who do not have Microsoft Excel installed. | Embed the generated HTML in a knowledge‑base or portal where readers can switch between sheet sections without leaving the page.
// AI Prompts: Write C# code with Aspose.Cells to export a multi‑sheet workbook to a single HTML file that includes a navigation pane listing all worksheets. | Explain the impact of ExportActiveWorksheetOnly, SaveAsSingleFile, and ShowAllSheets on the HTML output produced by Aspose.Cells. | Give troubleshooting steps when the navigation pane is missing after exporting an Excel workbook to HTML with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to build a workbook with several worksheets, fill them with data, and save the whole workbook as one HTML document that contains a navigation pane (sheet tabs) by setting HtmlSaveOptions: ExportActiveWorksheetOnly = false, SaveAsSingleFile = true, ShowAllSheets = true.
    public class ExportWithNavigationPane
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample worksheets
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Name = "Summary";
                Worksheet sheet1 = workbook.Worksheets.Add("Sales");
                Worksheet sheet2 = workbook.Worksheets.Add("Inventory");

                // Populate each worksheet with some data
                workbook.Worksheets["Summary"].Cells["A1"].PutValue("Company Overview");
                sheet1.Cells["A1"].PutValue("Product");
                sheet1.Cells["B1"].PutValue("Quantity");
                sheet1.Cells["A2"].PutValue("Widget");
                sheet1.Cells["B2"].PutValue(150);
                sheet2.Cells["A1"].PutValue("Item");
                sheet2.Cells["B1"].PutValue("Stock");
                sheet2.Cells["A2"].PutValue("Gadget");
                sheet2.Cells["B2"].PutValue(85);

                // Configure HTML save options to generate a single HTML file with a navigation pane
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export the whole workbook (all worksheets)
                    ExportActiveWorksheetOnly = false,
                    // Save everything into one HTML file
                    SaveAsSingleFile = true,
                    // Show all sheets as tabs in the navigation pane
                    ShowAllSheets = true
                };

                // Save the workbook as HTML; the resulting file contains a navigation pane linking to each sheet
                workbook.Save("WorkbookWithNavigation.html", saveOptions);
                Console.WriteLine("Workbook exported successfully to WorkbookWithNavigation.html");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWithNavigationPane.Run();
        }
    }
}

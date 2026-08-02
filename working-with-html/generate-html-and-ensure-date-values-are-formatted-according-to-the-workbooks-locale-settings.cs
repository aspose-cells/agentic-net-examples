// Title: Export Aspose.Cells Workbook to HTML5 with UK Locale Date Formatting (C#)
// Description: Demonstrates setting a workbook’s region to United Kingdom (dd/MM/yyyy), inserting a DateTime, applying locale‑aware styling, retrieving the cell’s HTML via GetHtmlString, and saving the entire workbook as HTML5 using HtmlSaveOptions.
// Keywords: Aspose.Cells | HTML5 export | UK locale | date formatting | GetHtmlString | HtmlSaveOptions | C# example | regional settings | CountryCode.UnitedKingdom | locale‑aware date
// Common Searches: Aspose.Cells export to HTML with UK date format | C# set workbook region United Kingdom Aspose | GetHtmlString locale date | HtmlSaveOptions HTML5 date formatting | How to format dates according to workbook locale in Aspose.Cells
// Developer Intent: Create HTML5 output from a workbook where dates follow the United Kingdom locale.
// Use Cases: Generate a financial statement web page with dd/MM/yyyy dates by configuring the workbook region before export. | Extract a single cell’s HTML snippet that automatically respects the workbook’s UK locale without manual string formatting. | Save a multi‑sheet report as HTML5 with consistent UK date formatting across all worksheets.
// AI Prompts: Write C# code to set Aspose.Cells workbook region to United Kingdom and export to HTML5 with locale‑aware dates. | Show how to use GetHtmlString on a date cell so the HTML reflects the workbook’s regional settings without custom formatting. | Explain how HtmlSaveOptions.HtmlVersion.Html5 and ExportDataOptions.All influence date rendering in HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlLocaleDemo
{
    // Demonstrates setting a workbook’s region to United Kingdom (dd/MM/yyyy), inserting a DateTime, applying locale‑aware styling, retrieving the cell’s HTML via GetHtmlString, and saving the entire workbook as HTML5 using HtmlSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Set the workbook's regional settings (locale) to United Kingdom (dd/MM/yyyy)
            workbook.Settings.Region = CountryCode.UnitedKingdom;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a DateTime value into cell A1
            Cell dateCell = sheet.Cells["A1"];
            dateCell.PutValue(DateTime.Now);

            // Apply a date format that respects the locale (optional, Aspose will use locale if no custom format)
            Style style = dateCell.GetStyle();
            style.Custom = "dd/MM/yyyy";
            dateCell.SetStyle(style);

            // Retrieve the HTML representation of the cell (including formatting)
            // The parameter 'true' adds a surrounding <div> wrapper, suitable for HTML5
            string cellHtml = dateCell.GetHtmlString(true);
            Console.WriteLine("Cell HTML:");
            Console.WriteLine(cellHtml);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Ensure formulas are calculated (not needed here but good practice)
                CalculateFormula = true,
                // Export all data (including styles) to HTML
                ExportDataOptions = HtmlExportDataOptions.All,
                // Use HTML5 standard
                HtmlVersion = HtmlVersion.Html5
            };

            // Save the entire workbook as an HTML file (lifecycle: save)
            workbook.Save("WorkbookWithLocaleDate.html", saveOptions);

            Console.WriteLine("HTML file saved. Date values are formatted according to the workbook's locale.");
        }
    }
}

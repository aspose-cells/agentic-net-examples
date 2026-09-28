// Title: Convert an HTML file to an Excel workbook using French (fr-FR) locale in C# with Aspose.Cells
// AI Prompts: Generate C# code that loads an HTML document into an Aspose.Cells Workbook, applies a French (fr-FR) CultureInfo to HtmlLoadOptions and Workbook settings, and saves the result as an XLSX file. | Show how to configure Aspose.Cells so that decimal separators are interpreted according to the French locale when converting HTML to Excel in .NET.
// Common Searches: aspocells html to xlsx using fr-FR cultureinfo | c# convert html table to excel with French decimal format | configure HtmlLoadOptions for French number formatting in Aspose.Cells | load html into workbook with French regional settings .NET | excel export from html respecting French number separators Aspose.Cells
// Tags: html to xlsx conversion respecting regional settings | Aspose.Cells load HTML with fr-FR culture | configure Workbook.Settings culture in C# | decimal separator handling Aspose.Cells | C# HtmlLoadOptions culture setup

using System;
using System.Globalization;
using Aspose.Cells;

// // Loads an HTML file into an Aspose.Cells Workbook with French (fr-FR) locale, sets the workbook's culture, and saves it as an XLSX file.
class Program
{
    static void Main()
    {
        // Input HTML file path
        string htmlFile = "input.html";

        // Output Excel file path
        string excelFile = "output.xlsx";

        // Create load options and set French culture (France) for proper decimal handling
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            CultureInfo = new CultureInfo("fr-FR")
        };

        // Load the HTML content into a new workbook using the specified culture
        Workbook workbook = new Workbook(htmlFile, loadOptions);

        // Ensure the workbook's culture is also set to French (affects further operations)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Save the workbook as an Excel file
        workbook.Save(excelFile, SaveFormat.Xlsx);
    }
}

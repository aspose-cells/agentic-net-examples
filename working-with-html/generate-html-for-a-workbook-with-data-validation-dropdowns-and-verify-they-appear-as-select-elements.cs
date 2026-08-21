// Title: Export Excel List Validation as HTML <select> using Aspose.Cells for .NET
// Description: Creates a workbook, adds a List‑type data‑validation to cell A1, enables the in‑cell dropdown, saves the sheet as HTML, and verifies that the output contains a <select> element.
// Keywords: Aspose.Cells | C# | HTML export | data validation list | in‑cell dropdown | select element | Excel to HTML conversion | validation rendering
// Common Searches: Aspose.Cells export data validation as dropdown | HTML <select> from Excel list validation C# | verify select tag in Aspose.Cells HTML output | how to render Excel dropdown in HTML with Aspose | C# generate HTML with data‑validation list
// Developer Intent: Generate an HTML file from an Excel workbook that preserves a List‑type validation as a functional <select> control and programmatically confirm its presence.
// Use Cases: Publish Excel worksheets with interactive dropdowns on web pages. | Automated regression tests to ensure data‑validation lists are rendered correctly in HTML exports. | Create static reports that retain user‑selectable options originally defined in Excel.
// AI Prompts: Write C# code with Aspose.Cells to add a List validation to cell B2, enable the in‑cell dropdown, export to HTML, and assert that a <select> tag exists in the output. | Explain how HtmlSaveOptions and HtmlExportDataOptions affect the rendering of Excel data‑validation lists as <select> elements in the generated HTML. | Provide a step‑by‑step guide to verify that an HTML file produced by Aspose.Cells includes the expected <select> element for a given validation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, adds a List‑type data‑validation to cell A1, enables the in‑cell dropdown, saves the sheet as HTML, and verifies that the output contains a <select> element.
class DataValidationHtmlDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define a cell area for the validation (cell A1)
        CellArea area = CellArea.CreateCellArea(0, 0, 0, 0);

        // Add a validation to the worksheet
        ValidationCollection validations = sheet.Validations;
        int index = validations.Add(area);
        Validation validation = validations[index];

        // Set validation type to List and provide comma‑separated values
        validation.Type = ValidationType.List;
        validation.Formula1 = "Option1,Option2,Option3";

        // Enable the in‑cell dropdown so it will be rendered as a <select> element in HTML
        validation.InCellDropDown = true;

        // Save the workbook to HTML
        string htmlPath = "DataValidationDemo.html";
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportDataOptions = HtmlExportDataOptions.All; // export full data
        workbook.Save(htmlPath, htmlOptions);

        // Verify that the generated HTML contains a <select> element
        string htmlContent = File.ReadAllText(htmlPath);
        bool containsSelect = htmlContent.IndexOf("<select", StringComparison.OrdinalIgnoreCase) >= 0;

        Console.WriteLine("HTML file generated at: " + Path.GetFullPath(htmlPath));
        Console.WriteLine("Contains <select> element for dropdown: " + (containsSelect ? "Yes" : "No"));
    }
}

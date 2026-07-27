// Title: Export Excel Validation List to HTML <select> with Aspose.Cells for .NET
// Description: Creates a workbook, adds a List‑type data‑validation to cell A1, enables the in‑cell dropdown so it renders as a <select> element, saves the sheet as HTML5 using HtmlSaveOptions, and verifies the generated HTML contains the <select> tag.
// Keywords: Aspose.Cells | .NET | C# | HTML export | Excel to HTML | data validation | list validation | dropdown list | select element | HtmlSaveOptions | Html5 | in‑cell dropdown | verification script
// Common Searches: Aspose.Cells export validation list as HTML select | C# generate HTML5 from Excel with dropdowns | How to render Excel data‑validation as <select> using Aspose.Cells | Check for <select> tag in Aspose.Cells HTML output | HtmlSaveOptions enable in‑cell dropdown Aspose.Cells
// Developer Intent: Generate HTML5 from an Excel workbook that includes a list validation rendered as a <select> control and programmatically confirm its presence.
// Use Cases: Embedding interactive Excel drop‑down lists in web pages without JavaScript libraries. | Automated testing of Aspose.Cells HTML export to ensure validation UI is preserved. | Creating printable HTML reports that retain user‑editable list selections.
// AI Prompts: Write C# code using Aspose.Cells to add a List validation to cell A1, enable InCellDropDown, export to HTML5, and assert that the output contains a <select> element. | Provide a step‑by‑step tutorial for verifying that Excel data‑validation lists appear as <select> tags after HTML export with Aspose.Cells. | Explain how to configure HtmlSaveOptions to include all data options and ensure validation dropdowns are rendered correctly in the generated HTML.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsValidationHtmlDemo
{
    // Creates a workbook, adds a List‑type data‑validation to cell A1, enables the in‑cell dropdown so it renders as a <select> element, saves the sheet as HTML5 using HtmlSaveOptions, and verifies the generated HTML contains the <select> tag.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a cell area for the validation (cell A1)
            CellArea area = CellArea.CreateCellArea(0, 0, 0, 0);

            // Add a validation to the worksheet
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(area);
            Validation validation = validations[validationIndex];

            // Set validation type to List and provide comma‑separated values
            validation.Type = ValidationType.List;
            validation.Formula1 = "Option1,Option2,Option3";

            // Enable the in‑cell dropdown so it will be rendered as a <select> element in HTML
            validation.InCellDropDown = true;

            // Save the workbook to HTML using HtmlSaveOptions
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDataOptions = HtmlExportDataOptions.All,
                HtmlVersion = HtmlVersion.Html5
            };
            string htmlPath = "ValidationDropdown.html";
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the generated HTML contains a <select> element
            string htmlContent = File.ReadAllText(htmlPath);
            bool containsSelect = htmlContent.IndexOf("<select", StringComparison.OrdinalIgnoreCase) >= 0;

            Console.WriteLine(containsSelect
                ? "Success: <select> element found in the generated HTML."
                : "Failure: <select> element not found in the generated HTML.");
        }
    }
}

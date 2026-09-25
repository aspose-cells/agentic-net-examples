// Title: Add a list‑type data‑validation dropdown to a cell range and export it as static text in a PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Create a list validation for cells B2:B10, set the dropdown items to "Apple,Banana,Cherry", and save the workbook as a PDF where the dropdown appears as plain text using Aspose.Cells in C#. | Load dropdown items from another worksheet column, apply the list validation to a target range, and generate a PDF that shows the selected values as static text. | Define multiple validation ranges, each with its own list of items, and export the workbook to PDF ensuring all dropdowns are rendered as static text.
// Common Searches: Aspose.Cells C# export Excel dropdown list as static text in PDF | How to preserve data validation list values when converting Excel to PDF with Aspose.Cells | Render in‑cell dropdown as plain text in PDF using Aspose.Cells for .NET
// Tags: list validation PDF export Aspose.Cells | cell range data validation C# Aspose.Cells | static rendering of dropdown in PDF | Aspose.Cells export Excel to PDF with validation | in‑cell dropdown as plain text PDF

using Aspose.Cells;
using System;

// Demonstrates adding a list‑type data‑validation dropdown to a specified cell range, setting initial values, and saving the workbook as a PDF where the dropdown is rendered as static text using Aspose.Cells for .NET (C#).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell range where the data‑validation dropdown will appear (e.g., B2:B10)
            CellArea area = new CellArea
            {
                StartRow = 1,    // Row index is zero‑based (1 = row 2)
                EndRow = 9,      // Row 10
                StartColumn = 1, // Column B
                EndColumn = 1
            };

            // Add a list‑type validation to the defined range
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;          // List dropdown
            validation.InCellDropDown = true;               // Show the dropdown arrow in Excel
            // Provide the list items as a comma‑separated string enclosed in double quotes
            validation.Formula1 = "\"Apple,Banana,Cherry\"";

            // Set an initial value for demonstration (optional)
            sheet.Cells["B2"].PutValue("Apple");

            // Save the workbook as PDF; the dropdown will be rendered as static text in the PDF
            workbook.Save("DataValidationDropdown.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

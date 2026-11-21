// Title: Export an Excel worksheet to PDF with data‑validation list values displayed as static text using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to PDF showing data validation list as plain text | How to render a dropdown validation cell as static text in a PDF with Aspose.Cells .NET | Save Excel worksheet with hidden list sheet to PDF while preserving default dropdown value in C# | Convert Excel cell that has list validation to PDF static value using Aspose.Cells | C# generate PDF from workbook where data‑validation dropdown appears as text
// Tags: PDF export data validation list Aspose.Cells | static text rendering for dropdown validation C# | hidden worksheet for list values Aspose.Cells | list‑type validation to PDF conversion .NET | Excel to PDF static dropdown value

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds a hidden sheet with the allowed list items, applies a list‑type validation to cell A1, sets a default value, and then saves the sheet as a PDF. The resulting PDF shows the selected dropdown value as static text instead of an interactive dropdown.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the list of allowed values for the dropdown
            string[] dropdownItems = new string[] { "Option1", "Option2", "Option3" };

            // Add a hidden sheet to store the list values
            int listSheetIndex = workbook.Worksheets.Add();
            Worksheet listSheet = workbook.Worksheets[listSheetIndex];
            listSheet.Name = "Lists";

            // Populate the list values in column A of the hidden sheet
            for (int i = 0; i < dropdownItems.Length; i++)
            {
                listSheet.Cells[i, 0].PutValue(dropdownItems[i]); // A1, A2, A3...
            }

            // Hide the sheet that contains the list (optional)
            listSheet.IsVisible = false;

            // Define the cell range for the validation (A1)
            CellArea validationArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 0,
                EndColumn = 0
            };

            // Add a validation (dropdown) to cell A1 on the main sheet
            ValidationCollection validations = sheet.Validations;
            int validationIndex = validations.Add(validationArea); // use overload with CellArea
            Validation validation = validations[validationIndex];

            validation.Type = ValidationType.List;
            // Reference the list range on the hidden sheet
            validation.Formula1 = $"=Lists!$A$1:$A${dropdownItems.Length}";
            // Show the dropdown arrow in the cell
            validation.InCellDropDown = true;

            // Set a default value that will be displayed in the PDF (static text)
            sheet.Cells["A1"].PutValue(dropdownItems[0]);

            // Determine output path and ensure the directory exists
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DataValidationDropdown.pdf");
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

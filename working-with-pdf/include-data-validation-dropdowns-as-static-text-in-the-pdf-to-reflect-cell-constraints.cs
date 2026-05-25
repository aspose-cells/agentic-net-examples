using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell area for the validation (cell A1)
            CellArea validationArea = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 0,
                EndColumn = 0
            };

            // Add a validation to the worksheet for the defined area
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation as a list with a dropdown
            validation.Type = ValidationType.List;
            validation.Formula1 = "Option1,Option2,Option3";
            validation.InCellDropDown = true; // Enable the in‑cell dropdown

            // Set PDF save options (property RenderDataValidationAsStaticText not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save("DataValidationStatic.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
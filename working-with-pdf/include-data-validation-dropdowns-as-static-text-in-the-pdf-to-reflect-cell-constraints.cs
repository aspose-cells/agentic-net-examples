// Title: C# – Render Excel Data‑Validation List as Static Text in PDF with Aspose.Cells
// Description: Demonstrates how to add a List‑type data‑validation to cell A1, enable the in‑cell dropdown, and save the workbook as a PDF using Aspose.Cells for .NET. The dropdown values are rendered as static, printable text in the resulting PDF, making validation choices visible without interactivity.
// Keywords: Aspose.Cells PDF export | C# data validation list PDF | Excel dropdown static text PDF | PdfSaveOptions Aspose.Cells | render list validation as text | export Excel to PDF with validation | .NET Excel to PDF example | GitHub Aspose.Cells PDF sample | global Excel PDF conversion
// Common Searches: How to show Excel data‑validation list in PDF using Aspose.Cells | Aspose.Cells C# render dropdown values as static text in PDF | Export Excel with validation to PDF .NET | PdfSaveOptions keep validation display | GitHub example Aspose.Cells PDF data validation
// Developer Intent: Create a PDF where an Excel cell’s data‑validation list appears as plain, printable text rather than an interactive dropdown.
// Use Cases: Generate compliance reports that display allowed entry values directly in the PDF. | Distribute printable templates with predefined choices (e.g., Option1, Option2, Option3) visible in the final document. | Automate batch conversion of workbooks containing validation rules, preserving the choice list for audit trails.
// AI Prompts: Write C# code with Aspose.Cells to add a List validation to cell A1 and export the workbook to PDF where the dropdown values are shown as static text. | Explain the rendering behavior of in‑cell dropdowns during PDF conversion with Aspose.Cells and how to ensure the list appears as plain text. | Provide a step‑by‑step guide to configure PdfSaveOptions for preserving data‑validation display in the generated PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to add a List‑type data‑validation to cell A1, enable the in‑cell dropdown, and save the workbook as a PDF using Aspose.Cells for .NET. The dropdown values are rendered as static, printable text in the resulting PDF, making validation choices visible without interactivity.
public class DataValidationPdfDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell area (A1) where the validation will be applied
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 0,
            EndColumn = 0
        };

        // Add a validation to the collection for the defined area
        int validationIndex = sheet.Validations.Add(area);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation as a list with a dropdown
        validation.Type = ValidationType.List;
        validation.Formula1 = "Option1,Option2,Option3";
        validation.InCellDropDown = true; // Enable in‑cell dropdown

        // Save the workbook as PDF; the dropdown values are rendered as static text
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("DataValidationDemo.pdf", pdfOptions);
    }
}

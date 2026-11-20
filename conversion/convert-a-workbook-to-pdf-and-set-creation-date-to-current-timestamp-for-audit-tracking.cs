// Title: Convert an Excel workbook to PDF and set the creation date to the current time using Aspose.Cells for .NET
// Common Searches: c# aspnet convert xlsx to pdf and set document created time with Aspose.Cells | how to assign current timestamp to BuiltInDocumentProperties.CreatedTime before saving PDF in Aspose.Cells | Aspose.Cells example for loading workbook, updating metadata, and exporting to PDF | set audit creation date on Excel file when converting to PDF using Aspose.Cells C# | load password protected xlsx and save as pdf with updated CreatedTime using Aspose.Cells
// Tags: Aspose.Cells workbook to PDF conversion | set CreatedTime property Aspose.Cells | update BuiltInDocumentProperties timestamp C# | load workbook with LoadOptions Aspose.Cells | audit metadata during Excel to PDF export

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an XLSX workbook with Aspose.Cells, updates the BuiltInDocumentProperties.CreatedTime to the current system time for audit tracking, and saves the workbook as a PDF, handling both Aspose.Cells-specific and general exceptions.
class WorkbookToPdfConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // LoadOptions can be used to specify a password if needed.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // Uncomment and set the password if the workbook is protected.
            // loadOptions.Password = "yourPassword";

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set the creation date to the current timestamp for audit tracking.
            workbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., missing password).
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors.
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

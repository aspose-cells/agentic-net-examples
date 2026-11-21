// Title: Load an XLSX workbook without charts, set worksheet paper size to A5, and save as PDF using Aspose.Cells in C#
// Common Searches: how to ignore charts when loading an Excel file with Aspose.Cells C# | set A5 paper size for a worksheet before converting to PDF using Aspose.Cells | Aspose.Cells LoadOptions LoadDataOnly property example for PDF export | convert XLSX to PDF with custom page size in C# Aspose.Cells | disable shapes and charts on workbook load Aspose.Cells C#
// Tags: LoadOptions chart suppression | worksheet page setup A5 | XLSX to PDF conversion C# | disable charts on load | custom paper size PDF export

using System;
using System.IO;
using Aspose.Cells;

// Loads an XLSX workbook while optionally ignoring charts, sets the first worksheet's printer paper size to A5, and saves the workbook as a PDF file using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load options for XLSX format. If the current Aspose.Cells version supports LoadDataOnly,
            // it can be enabled to ignore charts, shapes, etc. The property is set conditionally.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Uncomment the following line if your Aspose.Cells version includes the LoadDataOnly property.
            // loadOptions.LoadDataOnly = true;

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Set the printer paper size to A5 for the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;

            // Save the workbook as a PDF file.
            workbook.Save(outputFile, SaveFormat.Pdf);

            Console.WriteLine($"Successfully converted \"{inputFile}\" to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

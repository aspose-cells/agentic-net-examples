// Title: How to load an Excel workbook with LoadOptions, set the first worksheet's printer paper size to Letter, and export it as PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook with LoadOptions then change page setup before PDF export | set worksheet paper size to Letter in Aspose.Cells before saving as PDF | verify generated PDF page count after setting paper size with Aspose.Cells C# | C# check if PDF file was created after converting Excel using Aspose.Cells | use PaperSizeType.PaperLetter in Aspose.Cells page setup
// Tags: LoadOptions workbook loading Aspose.Cells | worksheet PageSetup PaperSizeType.PaperLetter | custom page setup for Excel to PDF conversion | C# verify PDF generation Aspose.Cells | export Excel worksheet as PDF with specific paper size

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file (input.xlsx) using Aspose.Cells LoadOptions, changes the first worksheet's PageSetup.PaperSize to PaperLetter, saves the workbook as a PDF (output.pdf), and confirms that the PDF file was successfully created.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Verify that the input Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file '{excelPath}' was not found.");
            return;
        }

        try
        {
            // LoadOptions can be used to specify load behavior; here we just instantiate it
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified LoadOptions
            Workbook workbook = new Workbook(excelPath, loadOptions);

            // Set the printer paper size to Letter for the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;

            // Define the output PDF file path
            string pdfPath = "output.pdf";

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine($"Success: PDF file '{pdfPath}' was generated.");
            }
            else
            {
                Console.WriteLine($"Failure: PDF file '{pdfPath}' was not created.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

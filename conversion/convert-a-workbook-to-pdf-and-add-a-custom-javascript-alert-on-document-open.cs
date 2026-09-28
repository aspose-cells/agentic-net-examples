// Title: How to convert an Excel .xlsx workbook to PDF with Aspose.Cells in C# while verifying file existence
// AI Prompts: Generate C# code that checks for the presence of a specified .xlsx file, loads it with Aspose.Cells, and saves it as a PDF to a given output path. | Create a robust try‑catch block around the workbook loading and PDF saving steps, logging any exceptions to the console. | Modify the sample to accept the input Excel path and output PDF path as command‑line arguments and ensure the target directories exist before saving.
// Common Searches: asp.net convert excel file to pdf using aspose.cells with file existence validation | c# sample to load workbook and save as pdf handling errors | how to pass dynamic input and output paths to aspose.cells pdf conversion | asp.net core check if xlsx exists before converting to pdf with aspose | c# console app convert xlsx to pdf and log conversion status
// Tags: Aspose.Cells Excel to PDF conversion C# | file existence check before workbook conversion | exception handling for Aspose.Cells Save operation | dynamic input and output paths in Aspose.Cells PDF export | console logging of Aspose.Cells conversion result

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, converts the workbook to PDF using SaveFormat.Pdf, writes the result to 'output.pdf', and wraps the entire process in a try‑catch block that reports success or any errors to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input Excel file
            string excelPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(excelPath);

            // Define the output PDF path
            string pdfPath = "output.pdf";

            // Convert the workbook to PDF and save
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

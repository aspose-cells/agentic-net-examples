// Title: Bulk convert Excel workbooks to single‑page PDF using Aspose.Cells for .NET (FitToPagesWide=1, FitToPagesTall=1)
// AI Prompts: Write a C# console program that scans a folder for .xls, .xlsx, and .xlsm files, sets each worksheet's PageSetup FitToPagesWide and FitToPagesTall properties to 1, and saves the workbook as a PDF with Aspose.Cells. | Generate a .NET script that batch processes a directory of Excel files, applies a one‑page‑wide and one‑page‑tall print layout to every sheet, and exports each file to PDF using Aspose.Cells.
// Common Searches: aspnet bulk export excel to pdf with fit to page settings | c# set FitToPagesWide and FitToPagesTall for all worksheets before saving as pdf | how to batch convert xlsx files to single page pdf using Aspose.Cells | process multiple Excel workbooks and apply page scaling in Aspose.Cells .NET
// Tags: Aspose.Cells bulk Excel to PDF conversion | set worksheet FitToPagesWide FitToPagesTall C# | page setup scaling for PDF export Aspose.Cells | C# batch processing of Excel files Aspose.Cells | single-page PDF generation from Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program iterates through a specified input folder, loads each .xls/.xlsx/.xlsm workbook with Aspose.Cells, sets every worksheet's PageSetup FitToPagesWide and FitToPagesTall to 1 (forcing a single‑page layout), and saves the result as a PDF in an output directory.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputExcel";

        // Folder where the generated PDF files will be saved
        string outputFolder = @"C:\OutputPdf";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (xls, xlsx, xlsm) from the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string excelPath in excelFiles)
        {
            string extension = Path.GetExtension(excelPath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue; // Skip non‑Excel files

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(excelPath);

            // Set FitToPagesWide = 1 and FitToPagesTall = 1 for every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup setup = sheet.PageSetup;
                setup.FitToPagesWide = 1;
                setup.FitToPagesTall = 1;
            }

            // Determine the PDF file name and full path
            string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Export the workbook to PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
        }
    }
}

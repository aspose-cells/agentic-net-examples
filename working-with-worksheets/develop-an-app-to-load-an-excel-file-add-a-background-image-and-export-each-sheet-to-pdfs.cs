// Title: Add Background Image to Worksheets and Export Each Sheet as PDF – Aspose.Cells C# Example
// Description: C# program that loads an Excel workbook, applies a background image to every worksheet, creates a single‑sheet workbook for each sheet, converts each to PDF with Aspose.Cells ConversionUtility, and saves the PDFs to a designated folder.
// Keywords: Aspose.Cells background image worksheet | export Excel sheet to PDF C# | per‑sheet PDF conversion Aspose.Cells | single‑sheet workbook Aspose.Cells | C# Excel to PDF with watermark | Aspose.Cells ConversionUtility example
// Common Searches: set worksheet background image Aspose.Cells .NET | export each Excel worksheet to separate PDF Aspose | convert single sheet workbook to PDF using Aspose.Cells | add watermark image to Excel sheets before PDF conversion | C# code to batch convert Excel sheets to PDFs
// Developer Intent: Load an Excel file, attach a background image to every worksheet, and generate an individual PDF for each sheet using Aspose.Cells for .NET.
// Use Cases: Brand every page of a multi‑sheet report with a logo or watermark before distribution. | Create separate PDF invoices from a workbook where each sheet represents a client. | Automate batch conversion of large workbooks while preserving custom background graphics on each page.
// AI Prompts: Generate C# code that adds a background image to a worksheet and saves it directly as PDF without intermediate files. | Show how to use a PNG background image and control its layout (stretch, tile, center) on a worksheet with Aspose.Cells. | Explain techniques to stream large workbooks sheet‑by‑sheet to PDF in Aspose.Cells, minimizing memory consumption.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsSheetToPdf
{
    // C# program that loads an Excel workbook, applies a background image to every worksheet, creates a single‑sheet workbook for each sheet, converts each to PDF with Aspose.Cells ConversionUtility, and saves the PDFs to a designated folder.
    class Program
    {
        static void Main()
        {
            // Paths – adjust as needed
            string excelPath = "input.xlsx";          // source Excel file
            string backgroundImagePath = "bg.jpg";    // background image file
            string outputFolder = "PdfOutputs";       // folder for per‑sheet PDFs
            string tempFolder = "TempSheets";         // temporary folder for intermediate files

            try
            {
                // Verify required input files exist
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                if (!File.Exists(backgroundImagePath))
                {
                    Console.WriteLine($"Error: Background image not found at '{backgroundImagePath}'.");
                    return;
                }

                // Ensure output directories exist
                Directory.CreateDirectory(outputFolder);
                Directory.CreateDirectory(tempFolder);

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(excelPath);

                // Read background image into a byte array
                byte[] bgImageData = File.ReadAllBytes(backgroundImagePath);

                // Process each worksheet
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    try
                    {
                        Worksheet sheet = sourceWorkbook.Worksheets[i];

                        // Set background image for the current worksheet
                        sheet.BackgroundImage = bgImageData;

                        // Create a new workbook that will contain only this sheet
                        Workbook singleSheetWb = new Workbook();

                        // Remove the default empty sheet created by the constructor
                        singleSheetWb.Worksheets.Clear();

                        // Add a copy of the current sheet to the new workbook
                        singleSheetWb.Worksheets.AddCopy(sheet.Name);

                        // Define temporary Excel file name for this sheet
                        string tempExcelPath = Path.Combine(tempFolder, $"Sheet_{i}_{sheet.Name}.xlsx");

                        // Save the single‑sheet workbook to the temporary file
                        singleSheetWb.Save(tempExcelPath);

                        // Define the final PDF file name
                        string pdfPath = Path.Combine(outputFolder, $"{sheet.Name}.pdf");

                        // Convert the temporary Excel file to PDF using the provided ConversionUtility rule
                        ConversionUtility.Convert(tempExcelPath, pdfPath);

                        // Clean up the temporary Excel file
                        if (File.Exists(tempExcelPath))
                        {
                            File.Delete(tempExcelPath);
                        }

                        Console.WriteLine($"Worksheet '{sheet.Name}' exported to PDF: {pdfPath}");
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Error processing worksheet index {i}: {exSheet.Message}");
                    }
                }

                Console.WriteLine("All worksheets have been processed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}

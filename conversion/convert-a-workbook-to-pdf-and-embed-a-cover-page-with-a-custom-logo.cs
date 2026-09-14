// Title: Create a PDF from an Aspose.Cells workbook with a custom logo cover page in C#
// AI Prompts: Write C# code that builds an Excel workbook, inserts a logo image on the first worksheet as a cover page, styles a title cell, and saves the workbook as a PDF using Aspose.Cells. | Show how to add a picture to a worksheet, apply cell formatting, and configure PdfSaveOptions before exporting the workbook to PDF with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells add logo to first sheet and export as PDF with cover page | Aspose.Cells PDF export with custom cover image and title in .NET | Saving an Excel workbook to PDF while embedding a header logo using Aspose.Cells
// Tags: Aspose.Cells add picture to worksheet | Aspose.Cells export workbook to PDF | Aspose.Cells embed logo cover page | Aspose.Cells PdfSaveOptions configuration | C# generate PDF with cover sheet using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPdfWithCover
{
    // The program creates a new workbook, adds a cover worksheet with a logo image and styled title, populates a second worksheet with sample data, configures PDF save options, and saves the workbook as a PDF file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains a default worksheet)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare the cover page (first worksheet) with a logo
                // -------------------------------------------------
                Worksheet coverSheet = workbook.Worksheets[0];
                coverSheet.Name = "Cover";

                // Load the logo image if it exists
                string logoPath = "logo.png";
                if (File.Exists(logoPath))
                {
                    try
                    {
                        byte[] logoBytes = File.ReadAllBytes(logoPath);
                        using (MemoryStream ms = new MemoryStream(logoBytes))
                        {
                            // Add the logo picture to the cover sheet
                            // Parameters: upper left row, upper left column, lower right row, lower right column, image stream
                            int pictureIndex = coverSheet.Pictures.Add(1, 1, 10, 5, ms);
                            coverSheet.Pictures[pictureIndex].IsLocked = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Failed to insert logo. {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Logo file not found at '{logoPath}'. Skipping logo insertion.");
                }

                // Add a title below the logo
                var titleCell = coverSheet.Cells["A7"];
                titleCell.PutValue("Report Title");
                var titleStyle = titleCell.GetStyle();
                titleStyle.Font.IsBold = true;
                titleStyle.Font.Size = 24;
                titleCell.SetStyle(titleStyle);

                // -------------------------------------------------
                // 2. Add sample data to a second worksheet
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                dataSheet.Name = "Data";
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["A2"].PutValue("Apples");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["A3"].PutValue("Bananas");
                dataSheet.Cells["B3"].PutValue(85);
                dataSheet.Cells["A4"].PutValue("Cherries");
                dataSheet.Cells["B4"].PutValue(60);

                // -------------------------------------------------
                // 3. Configure PDF save options
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true,
                    CalculateFormula = true
                };

                // -------------------------------------------------
                // 4. Save the workbook as PDF
                // -------------------------------------------------
                string outputPdf = "WorkbookWithCover.pdf";
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"Workbook successfully saved to PDF: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

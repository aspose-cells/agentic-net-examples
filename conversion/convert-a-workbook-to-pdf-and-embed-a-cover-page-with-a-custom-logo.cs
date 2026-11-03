// Title: Create a PDF from an Excel workbook and prepend a custom logo cover page with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add a cover sheet with logo before exporting Excel to PDF | How to insert an image on the first page of a PDF generated from an Excel file using Aspose.Cells | C# combine a cover worksheet and existing workbook then save as PDF with Aspose.Cells | Add a title and logo to a PDF cover page when converting XLSX to PDF in .NET | Copy worksheets from one workbook to another and export as PDF with Aspose.Cells
// Tags: add cover worksheet with image Aspose.Cells | export workbook to PDF with custom cover Aspose.Cells | insert picture into worksheet C# Aspose.Cells | copy worksheets between workbooks Aspose.Cells | save combined workbook as PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel file, creates a new workbook for a cover sheet, inserts a logo image and title, copies all original worksheets into the new workbook, and saves the combined workbook as a PDF, resulting in a PDF whose first page is a custom cover with the logo.
class WorkbookToPdfWithCover
{
    static void Main()
    {
        // Paths for input and output files
        string sourceExcelPath = "input.xlsx";
        string logoImagePath = "logo.png";
        string outputPdfPath = "output.pdf";

        try
        {
            // Verify that required files exist
            if (!File.Exists(sourceExcelPath))
                throw new FileNotFoundException($"Source Excel file not found: {sourceExcelPath}");
            if (!File.Exists(logoImagePath))
                throw new FileNotFoundException($"Logo image file not found: {logoImagePath}");

            // Load the original workbook
            Workbook sourceWorkbook = new Workbook(sourceExcelPath);

            // Create a new workbook that will hold the cover page and the original content
            Workbook finalWorkbook = new Workbook();

            // Access the first worksheet of the new workbook – this will be the cover page
            Worksheet coverSheet = finalWorkbook.Worksheets[0];
            coverSheet.Name = "Cover";

            // Insert the logo image onto the cover sheet (position at cell A1)
            int pictureIndex = coverSheet.Pictures.Add(0, 0, logoImagePath);
            Picture logoPicture = coverSheet.Pictures[pictureIndex];
            logoPicture.Width = 300;
            logoPicture.Height = 100;

            // Add a title below the logo (cell A5)
            Cell titleCell = coverSheet.Cells["A5"];
            titleCell.PutValue("Report Title");
            Style titleStyle = titleCell.GetStyle();
            titleStyle.Font.Size = 24;
            titleStyle.Font.IsBold = true;
            titleCell.SetStyle(titleStyle);

            // Copy all worksheets from the source workbook into the final workbook after the cover sheet
            for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            {
                Worksheet sourceSheet = sourceWorkbook.Worksheets[i];
                // Add a copy of the source worksheet by name
                finalWorkbook.Worksheets.AddCopy(sourceSheet.Name);
            }

            // Save the combined workbook as PDF; the first page will be the cover with the logo
            finalWorkbook.Save(outputPdfPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF generated successfully at: {outputPdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

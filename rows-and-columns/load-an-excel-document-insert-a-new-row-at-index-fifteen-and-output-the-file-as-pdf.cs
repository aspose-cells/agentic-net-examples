// Title: Insert a row at index 15 in an Excel worksheet with Aspose.Cells for .NET and export the workbook to PDF
// AI Prompts: Load a .xlsx file with Aspose.Cells, insert a new row at row index 15 in the first worksheet, save the workbook, and convert it to a PDF using ConversionUtility. | Using Aspose.Cells for .NET, add a blank row at the 16th row of a worksheet and generate a PDF from the modified workbook.
// Common Searches: Aspose.Cells C# insert row at specific index and save as PDF | How to add a blank row at row 15 in Excel using Aspose.Cells .NET | Convert modified Excel workbook to PDF with Aspose.Cells ConversionUtility | C# example for inserting a row in worksheet and exporting to PDF using Aspose.Cells
// Tags: insert row Aspose.Cells API | Excel row insertion Aspose.Cells C# | workbook to PDF conversion Aspose.Cells | ConversionUtility PDF export Aspose.Cells | modify worksheet and generate PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads source.xlsx (creates a simple workbook if missing), inserts a blank row at index 15 in the first worksheet, saves the modified workbook, converts it to result.pdf using Aspose.Cells ConversionUtility, and deletes the temporary Excel file.
class InsertRowAndConvertToPdf
{
    static void Main()
    {
        // Define file paths
        string sourceExcelPath = @"C:\Input\source.xlsx";
        string modifiedExcelPath = @"C:\Input\source_modified.xlsx";
        string outputPdfPath = @"C:\Output\result.pdf";

        try
        {
            // Ensure the source file exists; create a simple workbook if it does not
            if (!File.Exists(sourceExcelPath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(sourceExcelPath));
                Workbook tempWb = new Workbook();
                Worksheet tempWs = tempWb.Worksheets[0];
                tempWs.Name = "Sheet1";
                tempWs.Cells["A1"].PutValue("Sample Data");
                tempWb.Save(sourceExcelPath);
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(sourceExcelPath);

            // Insert a new row at index 15 (16th row)
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells.InsertRow(15);

            // Ensure the directory for the modified file exists
            Directory.CreateDirectory(Path.GetDirectoryName(modifiedExcelPath));
            // Save the modified workbook
            workbook.Save(modifiedExcelPath);

            // Ensure the directory for the PDF output exists
            Directory.CreateDirectory(Path.GetDirectoryName(outputPdfPath));
            // Convert the modified workbook to PDF
            ConversionUtility.Convert(modifiedExcelPath, outputPdfPath);

            // Clean up the temporary modified Excel file
            if (File.Exists(modifiedExcelPath))
            {
                File.Delete(modifiedExcelPath);
            }

            Console.WriteLine("Row inserted and PDF generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

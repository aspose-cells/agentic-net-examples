using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells P15:Q17
        // P = column 15 (zero‑based), row 15 = index 14
        // Total rows = 3 (15,16,17), total columns = 2 (P,Q)
        worksheet.Cells.Merge(14, 15, 3, 2);

        // Create a style with the desired date format
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "mm-dd-yyyy";

        // Apply the style to the merged cell (upper‑left cell of the range)
        worksheet.Cells["P15"].SetStyle(dateStyle);

        // Save the workbook to a temporary XLSX file (required for conversion)
        string tempXlsxPath = Path.Combine(Path.GetTempPath(), "tempWorkbook.xlsx");
        workbook.Save(tempXlsxPath);

        // Convert the temporary XLSX file to PDF using the provided ConversionUtility rule
        string pdfPath = "MergedCellsDateFormat.pdf";
        ConversionUtility.Convert(tempXlsxPath, pdfPath);

        // Clean up the temporary file
        if (File.Exists(tempXlsxPath))
        {
            File.Delete(tempXlsxPath);
        }

        Console.WriteLine($"PDF generated at: {Path.GetFullPath(pdfPath)}");
    }
}
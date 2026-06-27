using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJapaneseEraPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Set the regional settings to Japan so that date formatting uses Japanese conventions
            workbook.Settings.Region = CountryCode.Japan;

            // Access the first worksheet and a target cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Put a numeric value that represents a date (Excel serial date)
            // Example: 44089 corresponds to 2020-09-15
            cell.PutValue(44089);

            // Create a style that formats the cell using Japanese era format
            // The custom format uses the locale code for Japanese ([$-F800])
            // and displays year, month, and day.
            Style style = cell.GetStyle();
            style.Custom = "[$-F800]yyyy年m月d日";
            cell.SetStyle(style);

            // Save the workbook to a temporary Excel file (required for conversion)
            string tempExcelPath = "temp.xlsx";
            workbook.Save(tempExcelPath, SaveFormat.Xlsx);

            // Convert the saved Excel file to PDF using the provided ConversionUtility rule
            string pdfPath = "JapaneseEraDate.pdf";
            ConversionUtility.Convert(tempExcelPath, pdfPath);

            Console.WriteLine("PDF generated successfully at: " + pdfPath);
        }
    }
}
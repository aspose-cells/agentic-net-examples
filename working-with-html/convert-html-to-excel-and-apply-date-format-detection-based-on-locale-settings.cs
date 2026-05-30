using System;
using System.Globalization;
using Aspose.Cells;

namespace HtmlToExcelWithLocaleDateDetection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input HTML file path
            string htmlPath = "input.html";

            // Output Excel file path
            string excelPath = "output.xlsx";

            // Locale identifier (e.g., "en-US", "fr-FR")
            string locale = "en-US";

            // Load the HTML file into a workbook
            // Aspose.Cells can directly load HTML documents
            Workbook workbook = new Workbook(htmlPath);

            // Get the first worksheet (or iterate all worksheets if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare culture info for date parsing based on the provided locale
            CultureInfo culture = new CultureInfo(locale);

            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue.Trim();

                        // Try to parse the string as a DateTime using the specified locale
                        if (DateTime.TryParse(text, culture, DateTimeStyles.None, out DateTime dt))
                        {
                            // Replace the string with an actual DateTime value
                            cell.PutValue(dt);

                            // Apply a standard date number format (e.g., "mm/dd/yyyy")
                            Style style = cell.GetStyle();
                            style.Number = 14; // Built‑in date format
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the workbook as an Excel file
            workbook.Save(excelPath);
        }
    }
}
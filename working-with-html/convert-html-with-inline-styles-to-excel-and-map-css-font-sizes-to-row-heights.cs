using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelWithRowHeight
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file that contains inline styles
            string htmlPath = "input.html";

            // Verify that the HTML file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file \"{htmlPath}\" was not found.");
                return;
            }

            try
            {
                // Load the HTML file into a workbook using default HtmlLoadOptions
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Get the first worksheet (adjust if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Iterate through all rows that contain data
                int maxRow = sheet.Cells.MaxDataRow;
                for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
                {
                    double maxFontSize = 0.0;

                    // Examine each cell in the current row to find the largest font size
                    int maxCol = sheet.Cells.MaxDataColumn;
                    for (int colIndex = 0; colIndex <= maxCol; colIndex++)
                    {
                        Cell cell = sheet.Cells[rowIndex, colIndex];
                        if (cell != null && cell.Value != null)
                        {
                            // Retrieve the font size from the cell's style
                            double fontSize = cell.GetStyle().Font.Size;
                            if (fontSize > maxFontSize)
                            {
                                maxFontSize = fontSize;
                            }
                        }
                    }

                    // If a font size was found, map it to a row height.
                    // Aspose.Cells uses points for font size and points for row height.
                    // A simple conversion factor (e.g., 1.2) gives a comfortable spacing.
                    if (maxFontSize > 0)
                    {
                        // Access the row via the Cells.Row collection
                        Row row = sheet.Cells.Rows[rowIndex];
                        row.Height = maxFontSize * 1.2;
                    }
                }

                // Save the workbook as an Excel file
                string excelPath = "output.xlsx";
                workbook.Save(excelPath, SaveFormat.Xlsx);

                Console.WriteLine("HTML converted to Excel with row heights adjusted based on CSS font sizes.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
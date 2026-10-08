// Title: Convert HTML to XLSX with Aspose.Cells in C# while detecting and formatting dates using a specific locale
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook, sets CultureInfo to a chosen locale, iterates through every used cell, converts recognizable date strings to true DateTime values, applies the built‑in short date number format, and saves the result as an XLSX file. | Demonstrate how to configure HtmlLoadOptions and Workbook.Settings.CultureInfo so that Aspose.Cells automatically parses dates from imported HTML according to US English (or any other) regional settings. | Provide robust error‑handling snippets that verify the source HTML file exists, create the destination folder if missing, and gracefully report conversion failures.
// Common Searches: c# aspose.cells convert html to excel with locale specific date parsing | how to detect and format dates when loading html into a workbook using asp.net | set cultureinfo for date detection in aspose.cells html import | auto apply short date format after importing html with aspose.cells c# | asp.net load html file into workbook and replace string dates with datetime cells
// Tags: HTML to XLSX conversion using Aspose.Cells | locale based date detection in Aspose.Cells workbook | Workbook.Settings.CultureInfo for date parsing | auto replace string dates with DateTime cells | apply built‑in short date format programmatically

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an HTML document into an Aspose.Cells Workbook, sets the workbook's CultureInfo (e.g., en-US) to guide date parsing, scans each used cell for string values that match date patterns, converts those strings to true DateTime cells, applies the built‑in short date number format, and saves the workbook as an XLSX file. It also includes checks for the input file's existence and creates the output directory when needed.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Input HTML file path
        string htmlPath = @"C:\Input\sample.html";

        // Output Excel file path
        string excelPath = @"C:\Output\result.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {htmlPath}");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the HTML file into a new Workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Set the locale for date detection (e.g., United States)
            workbook.Settings.CultureInfo = CultureInfo.GetCultureInfo("en-US");

            // Iterate through all worksheets and cells to detect dates
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Loop through each cell in the used range
                for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
                {
                    for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only cells that contain string values
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue.Trim();

                            // Try to parse the string as a DateTime using the workbook's culture
                            if (DateTime.TryParse(text, workbook.Settings.CultureInfo, DateTimeStyles.None, out DateTime parsedDate))
                            {
                                // Replace the string with a true date value
                                cell.PutValue(parsedDate);

                                // Apply a date number format (short date)
                                Style style = cell.GetStyle();
                                style.Number = 14; // Built‑in short date format
                                cell.SetStyle(style);
                            }
                        }
                    }
                }
            }

            // Save the workbook as an Excel file
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Conversion completed successfully. Excel saved to: {excelPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}

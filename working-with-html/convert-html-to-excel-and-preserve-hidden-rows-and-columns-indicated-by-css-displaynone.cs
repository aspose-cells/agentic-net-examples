using System;
using System.IO;
using Aspose.Cells;

class HtmlToExcelConverter
{
    static void Main()
    {
        // Input HTML file that may contain CSS display:none for rows/columns
        string htmlFile = "input.html";

        // Output Excel file where hidden rows/columns will be preserved
        string excelFile = "output.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlFile))
                throw new FileNotFoundException($"Input file not found: {htmlFile}");

            // Load the HTML document into a workbook.
            // Aspose.Cells interprets CSS display:none as hidden rows/columns during load.
            Workbook workbook = new Workbook(htmlFile);

            // Save the workbook in XLSX format.
            workbook.Save(excelFile, SaveFormat.Xlsx);

            Console.WriteLine("HTML has been converted to Excel with hidden rows and columns preserved.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
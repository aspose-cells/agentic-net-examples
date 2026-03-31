using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Paths for the source HTML file and the destination XLSX file
        string htmlFilePath = "input.html";
        string xlsxFilePath = "output.xlsx";

        // Configure HTML load options
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            // Import formulas present in the HTML (if any)
            LoadFormulas = true
        };

        // Convert HTML tables to Excel ListObjects for better structure
        loadOptions.TableLoadOptions.TableToListObject = true;

        // Load the HTML content into a workbook using the specified options
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Save the workbook as an XLSX file
        workbook.Save(xlsxFilePath);
    }
}
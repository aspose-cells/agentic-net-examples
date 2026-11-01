// Title: Import an HTML table into an Excel worksheet and convert it to a structured ListObject using Aspose.Cells for .NET
// AI Prompts: Load an HTML file containing a table into a Workbook, determine the occupied cell range, and add a ListObject that uses the first row as headers. | Give the ListObject a custom display name, apply a built‑in table style, and save the workbook as an XLSX file.
// Common Searches: Aspose.Cells C# import HTML table and convert to Excel ListObject | How to create a structured table from an HTML file with Aspose.Cells .NET | Set display name and style for a ListObject after loading HTML in Aspose.Cells | Determine min and max row/column after loading HTML into a workbook using Aspose.Cells
// Tags: html to workbook conversion Aspose.Cells | add ListObject from cell range Aspose.Cells | assign display name to ListObject Aspose.Cells | apply built‑in table style Aspose.Cells | export workbook to xlsx Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an HTML file that contains a table into an Aspose.Cells Workbook, identifies the data range, creates a ListObject (structured table) over that range with a header row, assigns a display name and a built‑in style, and saves the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the HTML file that contains the table
            string htmlPath = "table.html";

            // Ensure the HTML file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file \"{htmlPath}\" was not found.");
                return;
            }

            // Load the HTML file into a new workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Get the first worksheet where the HTML table was placed
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the range occupied by the imported data
            int firstRow = sheet.Cells.MinRow;
            int firstColumn = sheet.Cells.MinColumn;
            int lastRow = sheet.Cells.MaxRow;
            int lastColumn = sheet.Cells.MaxColumn;

            // Add a ListObject (structured table) to the determined range.
            // The method returns the index of the newly added ListObject.
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table (Name property does not exist)
            table.DisplayName = "ImportedHtmlTable";

            // Optional: apply a built‑in table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook as an Excel file
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully as Result.xlsx.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

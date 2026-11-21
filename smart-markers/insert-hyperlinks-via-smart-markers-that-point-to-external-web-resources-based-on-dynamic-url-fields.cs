// Title: Create clickable external hyperlinks in an Excel worksheet from a DataTable using Aspose.Cells HYPERLINK formula in C#
// Common Searches: asp.net generate Excel file with clickable links from a DataTable using Aspose.Cells | c# add HYPERLINK formula to Excel cells from a list of URLs with Aspose.Cells | how to export URLs as hyperlinked cells in Excel using Aspose.Cells C# | create Excel workbook with dynamic external hyperlinks from DataTable in a console application | save Excel workbook with "Click Here" hyperlinks using Aspose.Cells library
// Tags: Aspose.Cells insert HYPERLINK formula | C# generate Excel hyperlinks from DataTable | Excel external link creation with Aspose | save workbook with clickable URLs Aspose.Cells | populate worksheet cells with URLs using Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, builds a DataTable of URL strings, and iterates through each row to assign an Excel HYPERLINK formula (display text "Click Here") to successive cells in the first column. The workbook is saved as Hyperlinks.xlsx, and the full file path is printed, with exception handling for robustness.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a data source containing dynamic URLs.
            DataTable urlTable = new DataTable("Links");
            urlTable.Columns.Add("url", typeof(string));
            urlTable.Rows.Add("https://www.example.com");
            urlTable.Rows.Add("https://www.openai.com");
            urlTable.Rows.Add("https://www.github.com");

            // Insert a hyperlink for each URL in the data table.
            // The displayed text for each hyperlink is set to "Click Here".
            for (int i = 0; i < urlTable.Rows.Count; i++)
            {
                string url = urlTable.Rows[i]["url"].ToString();
                // Use the HYPERLINK formula to create a clickable link.
                sheet.Cells[i, 0].Formula = $"=HYPERLINK(\"{url}\", \"Click Here\")";
            }

            // Define output file path.
            string outputPath = "Hyperlinks.xlsx";

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

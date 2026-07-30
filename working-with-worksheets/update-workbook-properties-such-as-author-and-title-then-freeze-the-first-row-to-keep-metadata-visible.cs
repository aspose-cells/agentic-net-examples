// Title: Set Author & Title Properties and Freeze Top Row in Excel with Aspose.Cells for .NET
// Description: Creates a new workbook, assigns built‑in Author and Title metadata, freezes the first worksheet row using FreezePanes, ensures the output folder exists, and saves the file as UpdatedWorkbook.xlsx using C#.
// Keywords: Aspose.Cells C# set document properties | freeze first row Excel | Workbook BuiltInDocumentProperties | FreezePanes example | save workbook to folder
// Common Searches: how to add author and title to Excel file with Aspose.Cells | freeze header row Aspose.Cells .NET | Aspose.Cells FreezePanes syntax for rows only | save workbook after updating properties C#
// Developer Intent: Add author and title metadata to a workbook and keep the header row fixed while scrolling.
// Use Cases: Generate audit‑ready reports with author and title information embedded. | Produce large data sheets where the column headings remain visible during navigation. | Automate Excel file creation that includes metadata and a frozen top row before distribution.
// AI Prompts: Show how to set both built‑in and custom document properties in Aspose.Cells before saving. | Give an example of freezing multiple rows and columns together with FreezePanes in C#. | Explain how to programmatically confirm that the Author property was written correctly to the .xlsx file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates a new workbook, assigns built‑in Author and Title metadata, freezes the first worksheet row using FreezePanes, ensures the output folder exists, and saves the file as UpdatedWorkbook.xlsx using C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Update built‑in document properties
            workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Sample Workbook";

            // Freeze the first row (row index 0 is the header row, so freeze 1 row)
            Worksheet sheet = workbook.Worksheets[0];
            // FreezePanes(row, column, totalRows, totalColumns)
            // Freeze rows above row 1 (i.e., first row) and no columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Define output file path
            string outputPath = "UpdatedWorkbook.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to disk
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

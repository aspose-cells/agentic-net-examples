// Title: Export each worksheet of an Excel workbook to a separate TSV file with Aspose.Cells for .NET (C#)
// Common Searches: how to export every sheet of an Excel workbook to individual TSV files using Aspose.Cells in C# | Aspose.Cells C# code to save each worksheet as a tab‑separated values file | convert multi‑sheet Excel file to separate .tsv files programmatically | C# iterate Excel worksheets and write them to TSV with Aspose.Cells | export Excel worksheet data to TSV including column headers using Aspose.Cells | batch export Excel sheets to TSV format with Aspose.Cells .NET
// Tags: Aspose.Cells export worksheet to TSV | C# write DataTable as tab separated file | multiple Excel sheets to individual TSV files | convert Excel workbook sheets to TSV using Aspose.Cells | save Excel worksheet as .tsv in .NET

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, loops through each worksheet, extracts its used range into a DataTable (including column headers), and writes the data to a tab‑separated .tsv file named after the worksheet.
class ExcelToTsvConverter
{
    static void Main(string[] args)
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range dimensions
            int maxRow = sheet.Cells.MaxDataRow;
            int maxColumn = sheet.Cells.MaxDataColumn;

            // Export the worksheet data to a DataTable (including column names)
            DataTable dt = sheet.Cells.ExportDataTable(
                0,                     // start row
                0,                     // start column
                maxRow + 1,            // total rows
                maxColumn + 1,         // total columns
                true);                 // export column names as first row

            // Build the TSV file name based on the worksheet name
            string tsvFileName = $"{sheet.Name}.tsv";

            // Write the DataTable to a TSV file
            using (StreamWriter writer = new StreamWriter(tsvFileName))
            {
                foreach (DataRow row in dt.Rows)
                {
                    // Join each field with a tab character
                    string line = string.Join("\t", row.ItemArray);
                    writer.WriteLine(line);
                }
            }

            Console.WriteLine($"Worksheet '{sheet.Name}' exported to '{tsvFileName}'.");
        }
    }
}

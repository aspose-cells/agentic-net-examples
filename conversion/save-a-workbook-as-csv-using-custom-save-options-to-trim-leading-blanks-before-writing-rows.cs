// Title: Save a workbook as CSV in C# with Aspose.Cells while removing leading blank rows
// Common Searches: how to delete top empty rows in an Aspose.Cells worksheet before exporting to CSV in C# | C# Aspose.Cells save workbook as CSV using TxtSaveOptions and trim leading blanks | remove leading blank rows from a worksheet then export to CSV with Aspose.Cells | Aspose.Cells custom CSV save options UTF-8 without top blank rows | trim empty rows at the beginning of a sheet before CSV conversion using Aspose.Cells
// Tags: Aspose.Cells CSV export with TxtSaveOptions | remove leading blank rows worksheet Aspose.Cells | trim top empty rows before CSV save C# | custom CSV save options UTF-8 Aspose.Cells | delete first empty row Aspose.Cells API | worksheet row deletion Aspose.Cells C#

using System;
using System.Text;
using Aspose.Cells;

// The example creates a new Workbook, populates cells (including values with leading blanks), removes any completely empty rows at the top of the first worksheet, configures TxtSaveOptions for CSV format with UTF-8 encoding, and saves the result as a CSV file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells (example data)
            sheet.Cells["A1"].PutValue("  Hello");   // leading blanks in cell value
            sheet.Cells["A2"].PutValue("World");
            sheet.Cells["A3"].PutValue("");          // blank row (will be trimmed)

            // Trim leading blank rows manually (Aspose.Cells does not expose CsvSaveOptions in this version)
            TrimLeadingBlankRows(sheet);

            // Configure CSV save options using TxtSaveOptions
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
            {
                Encoding = Encoding.UTF8
            };

            // Save the workbook as a CSV file using the custom options
            workbook.Save("output.csv", csvOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Removes empty rows from the top of the worksheet
    private static void TrimLeadingBlankRows(Worksheet sheet)
    {
        // Continue deleting the first row while it is completely empty
        while (true)
        {
            bool isEmpty = true;
            // Determine the last column that contains data
            int maxCol = sheet.Cells.MaxDataColumn;
            // If there are no columns with data, the sheet is empty
            if (maxCol < 0)
                break;

            for (int col = 0; col <= maxCol; col++)
            {
                // Check if any cell in the first row has a non‑empty value
                if (!string.IsNullOrEmpty(sheet.Cells[0, col].StringValue))
                {
                    isEmpty = false;
                    break;
                }
            }

            if (isEmpty)
                sheet.Cells.DeleteRow(0);
            else
                break;
        }
    }
}

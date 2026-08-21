// Title: C# – Split Column B by Comma with Aspose.Cells TextToColumns
// Description: This example loads an Excel workbook, detects the populated rows in column B, configures TxtLoadOptions to use a comma separator, applies Cells.TextToColumns to expand each comma‑separated value into separate columns, and saves the result as a new file.
// Keywords: Aspose.Cells C# TextToColumns | Excel split column by comma | convert CSV cell to columns .NET | TxtLoadOptions separator | Aspose.Cells delimiter | C# Excel data parsing | TextToColumns example | Aspose.Cells workbook manipulation | comma delimiter Excel C# | split multi‑value cell Aspose
// Common Searches: Aspose.Cells separate values in column B using comma C# | How to use TextToColumns in Aspose.Cells | Convert CSV list in a single Excel cell to multiple columns .NET | C# code for TextToColumns with custom delimiter | Aspose.Cells delimiter options example
// Developer Intent: Expand each comma‑separated entry in column B into individual columns within an Excel file using Aspose.Cells.
// Use Cases: Transform tag lists stored in one cell into separate columns for reporting dashboards. | Prepare imported CSV strings for pivot tables by separating values into distinct columns. | Clean legacy spreadsheets that store multiple IDs in a single column before data migration. | Automate data normalization for SaaS platforms that receive multi‑value fields in Excel uploads.
// AI Prompts: Generate C# code that uses a semicolon as the delimiter instead of a comma for TextToColumns. | Provide a snippet that skips empty cells when applying TextToColumns to avoid exceptions. | Show how to apply TextToColumns to column D while automatically detecting the last row. | Create a version that writes the split data to a new worksheet instead of overwriting the original. | Explain how to chain TextToColumns with data validation to ensure numeric values after splitting.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example loads an Excel workbook, detects the populated rows in column B, configures TxtLoadOptions to use a comma separator, applies Cells.TextToColumns to expand each comma‑separated value into separate columns, and saves the result as a new file.
    class ConvertColumnBToSeparateColumns
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the number of rows that contain data in column B (index 1)
            int totalRows = sheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based

            // Configure text load options to use comma as the delimiter
            TxtLoadOptions options = new TxtLoadOptions();
            options.Separator = ',';               // Set comma as the separator
            options.TreatConsecutiveDelimitersAsOne = false; // Optional: keep default behavior

            // Split the text in column B starting from the first row (row index 0)
            // Parameters: start row, start column (1 for column B), total rows, options
            sheet.Cells.TextToColumns(0, 1, totalRows, options);

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}

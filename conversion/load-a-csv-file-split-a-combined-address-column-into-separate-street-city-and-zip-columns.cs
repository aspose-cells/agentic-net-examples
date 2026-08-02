// Title: C# – Split a Combined Address Column from CSV into Street, City, and ZIP with Aspose.Cells
// Description: Import a CSV file using Aspose.Cells, apply TxtLoadOptions and the TextToColumns method to separate the first column into street, city, and ZIP fields, and save the transformed data as an XLSX workbook.
// Keywords: Aspose.Cells | C# CSV import | TextToColumns | address parsing | split column | CSV to XLSX conversion | TxtLoadOptions | data transformation | Excel automation | .NET spreadsheet library
// Common Searches: Aspose.Cells split address column CSV | C# TextToColumns example for address fields | How to separate street city zip in Aspose.Cells | Convert CSV to Excel and split columns .NET | Parse combined address with Aspose.Cells
// Developer Intent: Load a CSV file, divide a combined address column into separate street, city, and ZIP columns, and export the result as an Excel workbook.
// Use Cases: Prepare mailing lists by converting raw CSV exports into Excel files with distinct address components for mail‑merge. | Enable reporting on customer locations by splitting address data into separate columns after importing CSV data. | Automate legacy data migration where address fields are concatenated, ensuring each part is stored in its own Excel column.
// AI Prompts: Write C# code that uses Aspose.Cells to import a CSV, split the first column into street, city, and ZIP using TextToColumns, and save the output as XLSX. | Explain the role of TxtLoadOptions when configuring the separator for TextToColumns in Aspose.Cells. | Suggest robust error‑handling for missing address columns or irregular delimiters when applying TextToColumns in a CSV import workflow.

using System;
using Aspose.Cells;

namespace AddressSplitExample
{
    // Import a CSV file using Aspose.Cells, apply TxtLoadOptions and the TextToColumns method to separate the first column into street, city, and ZIP fields, and save the transformed data as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Import the CSV data into the worksheet (lifecycle: load)
            // Using comma as the delimiter and converting numeric data where possible
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Determine the number of rows that contain data after import
            int totalRows = cells.MaxDataRow + 1;

            // Configure split options: use comma as the separator
            TxtLoadOptions splitOptions = new TxtLoadOptions();
            splitOptions.Separator = ',';

            // Split the combined address column (assumed to be column A, index 0)
            // into separate columns for street, city, and zip (lifecycle: transform)
            cells.TextToColumns(0, 0, totalRows, splitOptions);

            // Save the resulting workbook (lifecycle: save)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}

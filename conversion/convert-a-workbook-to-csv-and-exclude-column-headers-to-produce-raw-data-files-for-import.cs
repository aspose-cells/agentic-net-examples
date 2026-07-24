// Title: Convert Excel Workbook to Header‑less CSV with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file, removes the first worksheet row (assumed header), and saves the result directly as a CSV file, producing a raw data file ready for bulk import or downstream processing.
// Keywords: Aspose.Cells CSV export | C# remove header row | Excel to CSV without headers | SaveFormat.Csv Aspose | delete first row Aspose.Cells | raw data CSV generation | bulk import CSV | Aspose.Cells .NET conversion
// Common Searches: Aspose.Cells export CSV without header | C# convert Excel to CSV skip first row | How to delete header row before saving as CSV using Aspose.Cells | Save Excel as CSV raw data C# Aspose | Remove column titles when exporting to CSV with Aspose.Cells
// Developer Intent: Create a CSV file from an Excel workbook while excluding the header row.
// Use Cases: Loading data into a database that expects header‑less CSV | Feeding raw CSV into ETL pipelines that strip metadata | Generating input files for legacy systems that reject column names | Preparing data extracts for analytics tools that require plain values
// AI Prompts: Write C# code using Aspose.Cells to convert an .xlsx to CSV and omit the first row. | Show how to configure Aspose.Cells SaveOptions to export a worksheet without column headers. | Explain how to delete multiple header rows before saving a CSV with Aspose.Cells in .NET.

using System;
using Aspose.Cells;

// Loads an .xlsx file, removes the first worksheet row (assumed header), and saves the result directly as a CSV file, producing a raw data file ready for bulk import or downstream processing.
class ConvertWorkbookToCsvWithoutHeaders
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Path for the resulting CSV file (raw data without column headers)
        string destPath = "output.csv";

        // Load the workbook from the source file
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet (you can adjust the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove the first row which typically contains column headers
        // This operation modifies the worksheet in‑place
        worksheet.Cells.DeleteRow(0);

        // Save the modified workbook as CSV using the Save(string, SaveFormat) rule
        workbook.Save(destPath, SaveFormat.Csv);
    }
}

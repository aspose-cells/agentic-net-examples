// Title: How to split a combined address column into Street, City, and Zip columns while loading a CSV with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# split address column CSV into separate street city zip fields | Load CSV with Aspose.Cells and separate combined address into multiple columns | Insert new columns and parse address string using Aspose.Cells .NET | C# Aspose.Cells example to split address column after loading CSV file | How to extract street, city, zip from a single column in a CSV using Aspose.Cells
// Tags: address field decomposition Aspose.Cells | add columns after CSV load Aspose.Cells | combined address string parsing Aspose.Cells | export workbook to XLSX after CSV edit Aspose.Cells | LoadOptions for CSV files Aspose.Cells .NET

using System;
using Aspose.Cells;

// Demonstrates loading a CSV file with Aspose.Cells LoadOptions, inserting three new columns, splitting a combined "Street, City, Zip" address string into separate Street, City, and Zip columns, and saving the result as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Load the CSV file using LoadOptions for CSV format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook("input.csv", loadOptions);
        Worksheet sheet = workbook.Worksheets[0];

        // Index of the combined address column (e.g., column A = 0)
        int addressColIndex = 0;

        // Insert three new columns after the address column for Street, City, and Zip
        sheet.Cells.InsertColumns(addressColIndex + 1, 3);

        // Optional: set header names for the new columns
        sheet.Cells[0, addressColIndex + 1].PutValue("Street");
        sheet.Cells[0, addressColIndex + 2].PutValue("City");
        sheet.Cells[0, addressColIndex + 3].PutValue("Zip");

        // Determine the last row with data
        int lastRow = sheet.Cells.MaxDataRow;

        // Iterate over each data row (skip header row at index 0)
        for (int row = 1; row <= lastRow; row++)
        {
            // Read the combined address string
            string combinedAddress = sheet.Cells[row, addressColIndex].StringValue?.Trim();
            if (string.IsNullOrEmpty(combinedAddress))
                continue;

            // Expected format: "Street, City, Zip"
            string[] parts = combinedAddress.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            string street = parts.Length > 0 ? parts[0].Trim() : string.Empty;
            string city   = parts.Length > 1 ? parts[1].Trim() : string.Empty;
            string zip    = parts.Length > 2 ? parts[2].Trim() : string.Empty;

            // Write split values into the new columns
            sheet.Cells[row, addressColIndex + 1].PutValue(street);
            sheet.Cells[row, addressColIndex + 2].PutValue(city);
            sheet.Cells[row, addressColIndex + 3].PutValue(zip);
        }

        // Save the modified workbook (as Excel for better column handling)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}

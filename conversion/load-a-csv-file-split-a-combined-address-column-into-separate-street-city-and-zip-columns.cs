// Title: C# example: Load a CSV with Aspose.Cells, split a combined address column into Street, City, and Zip columns, and save the result
// Common Searches: how to split a combined address column into separate street, city, and zip fields using Aspose.Cells in C# | Aspose.Cells C# load CSV and insert new columns for address components | C# parse address string from CSV and write street, city, zip with Aspose.Cells | saving modified CSV after address splitting with Aspose.Cells .NET | example code for extracting street, city, zip from a CSV using Aspose.Cells LoadOptions
// Tags: load csv with Aspose.Cells LoadOptions | split address column using Aspose.Cells Cells API | insert columns in worksheet Aspose.Cells | save modified csv Aspose.Cells SaveFormat | parse street city zip C# string manipulation

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a CSV file into an Aspose.Cells Workbook, inserts three new columns for Street, City, and Zip after the combined address column, parses each address string (expected format "Street, City ZIP") into its components, writes the split values into the new columns, and saves the updated data as a new CSV file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";
            const string outputPath = "output.csv";

            // Verify that the input CSV file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file into a Workbook
            var loadOptions = new LoadOptions(LoadFormat.Csv);
            var workbook = new Workbook(inputPath, loadOptions);
            var worksheet = workbook.Worksheets[0];

            // Index of the combined address column (zero‑based). Adjust as needed.
            int addressColIndex = 2; // e.g., column C
            int headerRow = 0;       // assuming first row contains headers

            // Determine the last row with data
            int lastRow = worksheet.Cells.MaxDataRow;

            // Insert three new columns for Street, City, and Zip after the address column
            worksheet.Cells.InsertColumns(addressColIndex + 1, 3);

            // Set header names for the new columns
            worksheet.Cells[headerRow, addressColIndex + 1].PutValue("Street");
            worksheet.Cells[headerRow, addressColIndex + 2].PutValue("City");
            worksheet.Cells[headerRow, addressColIndex + 3].PutValue("Zip");

            // Process each data row and split the address
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string address = worksheet.Cells[row, addressColIndex].StringValue;
                string street = string.Empty;
                string city = string.Empty;
                string zip = string.Empty;

                if (!string.IsNullOrWhiteSpace(address))
                {
                    // Expected format: "Street, City ZIP"
                    // Split on the first comma
                    var parts = address.Split(new[] { ',' }, 2);
                    street = parts[0].Trim();

                    if (parts.Length > 1)
                    {
                        // Remaining part contains city and zip
                        string cityZipPart = parts[1].Trim();

                        // Find the last space to separate city and zip
                        int lastSpace = cityZipPart.LastIndexOf(' ');
                        if (lastSpace > -1)
                        {
                            city = cityZipPart.Substring(0, lastSpace).Trim();
                            zip = cityZipPart.Substring(lastSpace + 1).Trim();
                        }
                        else
                        {
                            city = cityZipPart; // No zip found
                        }
                    }
                }

                // Write split values into the new columns
                worksheet.Cells[row, addressColIndex + 1].PutValue(street);
                worksheet.Cells[row, addressColIndex + 2].PutValue(city);
                worksheet.Cells[row, addressColIndex + 3].PutValue(zip);
            }

            // Optional: remove the original combined address column
            // worksheet.Cells.DeleteColumn(addressColIndex);

            // Save the modified data to a new CSV file using SaveFormat.Csv
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Processing complete. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

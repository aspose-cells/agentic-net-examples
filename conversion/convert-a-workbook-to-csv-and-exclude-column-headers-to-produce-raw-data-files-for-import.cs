// Title: Convert an Excel workbook to UTF-8 CSV files per worksheet while skipping the header row using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx file, iterates through each worksheet, and saves each sheet as a separate UTF-8 encoded CSV file, excluding the first row (column headers) from the output. | Show how to configure TxtSaveOptions in Aspose.Cells to export raw data without header rows and apply it when saving multiple worksheets to CSV. | Create a reusable method that accepts a workbook path and an optional flag to omit headers, then generates one CSV per sheet using Aspose.Cells and returns the list of generated file names.
// Common Searches: aspocells c# export each sheet to csv without header row | how to save excel worksheets as separate csv files raw data asp.net | convert workbook to multiple utf-8 csv files using Aspose.Cells and skip first row | c# Aspose.Cells TxtSaveOptions hide column names when saving to csv | generate raw csv files from excel workbook for data import asp.net
// Tags: aspocells export worksheet to csv without headers | utf-8 csv generation aspocells c# | multiple csv files per excel sheet aspocells | txtsaveoptions raw data csv aspocells | c# convert workbook to csv raw import | skip first row csv export aspocells

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example loads an Excel workbook, iterates over every worksheet, creates a temporary workbook for each sheet, and saves it as a UTF-8 encoded CSV file named after the worksheet while omitting the first row to produce raw data suitable for import.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Loop through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Create a temporary workbook containing only the current worksheet
                    Workbook sheetWb = new Workbook();
                    sheetWb.Worksheets.Clear(); // Remove the default sheet

                    // Copy the current sheet into the temporary workbook
                    // Use the overload that accepts arrays of worksheets and names
                    sheetWb.Worksheets.AddCopy(new Worksheet[] { sheet }, new string[] { sheet.Name });

                    // Configure CSV save options using TxtSaveOptions
                    TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                    {
                        Encoding = Encoding.UTF8 // Use UTF-8 encoding for the CSV file
                    };

                    // Build the output CSV file name (one file per worksheet)
                    string csvFileName = $"{sheet.Name}.csv";

                    // Save the worksheet as a CSV file using the specified options
                    sheetWb.Save(csvFileName, csvOptions);
                    Console.WriteLine($"Worksheet '{sheet.Name}' saved as '{csvFileName}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save worksheet '{sheet.Name}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

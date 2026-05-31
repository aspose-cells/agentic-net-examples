using System;
using System.IO;
using Aspose.Cells;

class RefreshPivotTablesDemo
{
    static void Main()
    {
        const string inputPath = "InputWithPivot.xlsx";
        const string outputPath = "OutputRefreshed.xlsx";

        try
        {
            // Verify that the input workbook exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the data table and pivot tables
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that holds the source table (e.g., named "DataSheet")
            Worksheet dataSheet = workbook.Worksheets["DataSheet"];
            if (dataSheet == null)
            {
                Console.WriteLine("Worksheet 'DataSheet' not found in the workbook.");
                return;
            }

            // Update the underlying data of the table (example modifications)
            dataSheet.Cells["B2"].PutValue(1500);
            dataSheet.Cells["B3"].PutValue(2500);
            // Add more data updates as needed...

            // Refresh all pivot tables in the workbook so they reflect the changed data
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook with refreshed pivot tables
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
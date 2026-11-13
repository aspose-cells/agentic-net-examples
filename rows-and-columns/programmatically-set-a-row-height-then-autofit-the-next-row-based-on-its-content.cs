// Title: How to set a fixed row height and then auto‑fit the next row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set specific row height and auto‑fit another row | C# example to fix first row height and auto‑fit second row with Aspose.Cells | How to use AutoFitRow after defining row height in Aspose.Cells for .NET | Programmatically set row height to 30 points and auto‑fit the following row using Aspose.Cells | Aspose.Cells .NET tutorial for adjusting row height and auto‑fitting rows
// Tags: Aspose.Cells define row height C# | Aspose.Cells AutoFitRow method | Rows collection height property Aspose.Cells | auto‑fit row after fixed height Aspose.Cells | Excel row height manipulation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, writes long and short text to the second row, sets the first row's height to 30 points via the Rows collection, auto‑fits the second row with AutoFitRow(1), and saves the file as RowHeightAndAutoFit.xlsx.
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

            // Example data for the second row (index 1) to demonstrate auto‑fit
            sheet.Cells["A2"].PutValue("This is a long piece of text that should cause the row to expand when auto‑fitted.");
            sheet.Cells["B2"].PutValue("Short text");

            // Set the height of the first row (index 0) to a fixed value (e.g., 30 points)
            // Use the Cells.Rows collection to access the row
            sheet.Cells.Rows[0].Height = 30; // Height is measured in points

            // Auto‑fit the second row (index 1) based on its content
            // The overload with a single argument is available in all supported versions
            sheet.AutoFitRow(1);

            // Define output file path
            string outputPath = "RowHeightAndAutoFit.xlsx";

            // Ensure the directory exists (in case a relative path is used)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: How to auto‑fit Excel row height to the tallest cell content in pixels while looping rows using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set row height after AutoFitRow based on pixel measurement | Iterate through rows and apply dynamic row height in an existing workbook with Aspose.Cells | Retrieve calculated row height in points and assign it back in Aspose.Cells .NET | Adjust Excel row heights to fit wrapped text when processing rows programmatically | AutoFit each row and preserve the computed height using Aspose.Cells for .NET
// Tags: set row height points Aspose.Cells | calculate row height pixels .NET | iterate rows Excel workbook Aspose | dynamic row height adjustment Excel | preserve computed row height Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a workbook, loops through all used rows, calls AutoFitRow to let Aspose.Cells compute the required height, reads the height in points with GetRowHeight, re‑applies the same height using SetRowHeight, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the last row that contains data
            int lastRow = cells.MaxDataRow;

            // Enumerate each row in the used range
            for (int rowIndex = 0; rowIndex <= lastRow; rowIndex++)
            {
                // AutoFit the row to let Aspose.Cells calculate the required height
                worksheet.AutoFitRow(rowIndex);

                // Retrieve the calculated height in points
                double heightInPoints = cells.GetRowHeight(rowIndex);

                // Explicitly set the row height using the same point value (ensures consistency)
                cells.SetRowHeight(rowIndex, heightInPoints);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

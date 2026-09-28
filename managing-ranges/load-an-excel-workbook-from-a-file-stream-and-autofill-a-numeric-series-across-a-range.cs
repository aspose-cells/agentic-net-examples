// Title: Load an Excel workbook from a FileStream and autofill a numeric series in column A using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file via FileStream, writes the value 1 to cell A1, creates a source range for A1, defines a destination range A1:A10, applies AutoFill with AutoFillType.Series, and writes the result to a new file. | Show how to use Aspose.Cells Workbook, Worksheet, and Range objects to programmatically fill a column with a sequential number series after loading the workbook from a stream.
// Common Searches: aspnet c# load excel file from filestream aspose.cells autofill series | how to use AutoFillType.Series in Aspose.Cells to fill a column | c# example creating numeric series A1:A10 after opening workbook from stream | aspose.cells load workbook from filestream and generate sequential numbers in column
// Tags: filestream workbook import Aspose.Cells | autofill sequential numbers Aspose.Cells | range definition for autofill Aspose.Cells | export modified workbook Aspose.Cells | input file existence validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the presence of input.xlsx, loads it into an Aspose.Cells Workbook via a FileStream, writes 1 to cell A1, defines source and destination ranges (A1 and A1:A10), uses AutoFill with AutoFillType.Series to generate a numeric series down column A, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from a file stream
            using (FileStream fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                Workbook workbook = new Workbook(fs);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set the first cell of the series
                Cell startCell = sheet.Cells["A1"];
                startCell.PutValue(1);

                // Define the source (starting cell) and the destination range to fill
                Aspose.Cells.Range source = sheet.Cells.CreateRange("A1");
                Aspose.Cells.Range destination = sheet.Cells.CreateRange("A1:A10");

                // Autofill a numeric series down the column
                source.AutoFill(destination, AutoFillType.Series);

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

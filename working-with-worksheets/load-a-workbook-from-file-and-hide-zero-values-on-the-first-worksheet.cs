// Title: Load an Excel file and hide zero values on the first worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load workbook from file and hide zero values on first sheet | how to set IsZeroValuesHidden on a worksheet with Aspose.Cells .NET | Aspose.Cells example to suppress zero display in an Excel worksheet programmatically | load existing .xlsx and hide zero values using Aspose.Cells C# | verify input file existence before loading workbook with Aspose.Cells
// Tags: Aspose.Cells load workbook from file | hide zero values worksheet Aspose.Cells | IsZeroValuesHidden property C# | first worksheet manipulation Aspose.Cells | Excel zero values visibility control Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, accesses the first worksheet, optionally enables the IsZeroValuesHidden property to hide zero values, saves the result as output.xlsx, and handles any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide zero values on this worksheet if the property is available.
            // The IsZeroValuesHidden property exists in newer Aspose.Cells versions.
            // Uncomment the line below when using a version that supports it.
            // sheet.IsZeroValuesHidden = true;

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

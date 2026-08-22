// Title: Convert string‑based numeric cells to true numbers in every worksheet of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX file, iterates through all worksheets, calls ConvertStringToNumericValue on each sheet's Cells collection, and saves the workbook to a new file. | Describe the steps to batch‑convert numeric text values to actual numbers across all sheets of an Excel workbook with Aspose.Cells in C#.
// Common Searches: Aspose.Cells .NET convert numeric text to numbers in all worksheets | C# batch convert string numbers to numeric values in Excel workbook | How to use ConvertStringToNumericValue on every sheet with Aspose.Cells | Change text‑formatted numbers to real numbers in an XLSX file using Aspose.Cells C# | ConvertStringToNumericValue method example for whole workbook
// Tags: ConvertStringToNumericValue across worksheets | numeric text to number conversion Aspose.Cells | batch numeric conversion Excel .NET | Aspose.Cells workbook numeric string handling | save workbook after numeric conversion Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsStringToNumberConversion
{
    // Loads an Excel workbook, loops through each worksheet, applies ConvertStringToNumericValue to turn numeric strings into true numeric cells, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with actual file path)
            string inputPath = "input.xlsx";

            // Path to the destination workbook after conversion
            string outputPath = "output_converted.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Convert all string values that can be interpreted as numbers to true numeric values
                sheet.Cells.ConvertStringToNumericValue();
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
        }
    }
}

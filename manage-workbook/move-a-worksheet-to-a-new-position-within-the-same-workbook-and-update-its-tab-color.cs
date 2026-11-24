// Title: Reorder a worksheet to the first position and change its tab color with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells move worksheet to first tab and set tab color | How to change Excel sheet tab color after moving a sheet using Aspose.Cells | Programmatically reorder worksheets in a workbook with Aspose.Cells .NET | Aspose.Cells MoveTo method example for changing worksheet order and tab color | Set LightBlue tab color for a specific worksheet in Aspose.Cells C#
// Tags: Worksheet.MoveTo API Aspose.Cells | tab color assignment Aspose.Cells | worksheet ordering in Excel using Aspose.Cells | modify sheet tab appearance .NET | move sheet to index zero Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, selects the worksheet named "DataSheet" (or the first sheet if not found), moves it to the first position (index 0) using the Worksheet.MoveTo method, changes its tab color to LightBlue, and saves the workbook as a new file, with error handling for missing files and exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Try to get the worksheet named "DataSheet"
                Worksheet sheet = workbook.Worksheets["DataSheet"];

                // If not found, fall back to the first worksheet
                if (sheet == null)
                {
                    Console.WriteLine("Worksheet 'DataSheet' not found. Using the first worksheet instead.");
                    if (workbook.Worksheets.Count == 0)
                    {
                        Console.WriteLine("The workbook contains no worksheets.");
                        return;
                    }
                    sheet = workbook.Worksheets[0];
                }

                // Move the worksheet to the first position (index 0)
                sheet.MoveTo(0);

                // Set a new tab color for the moved worksheet
                sheet.TabColor = Color.LightBlue;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

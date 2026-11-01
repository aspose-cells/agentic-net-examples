// Title: Remove every slicer from all worksheets in an Excel file and export the workbook as a single PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, iterates through each worksheet, clears the SlicerCollection on every sheet, and saves the result as one consolidated PDF with Aspose.Cells. | Write a console application that verifies the source Excel file exists, clears all slicer collections across worksheets, and outputs a single PDF while handling possible exceptions. | Create a method called ClearSlicersAndExportPdf(string inputPath, string outputPath) that removes slicer objects from every worksheet and calls Workbook.Save with SaveFormat.Pdf.
// Common Searches: c# aspose.cells how to delete slicer controls from each worksheet before PDF conversion | export an Excel workbook to one PDF after removing slicer objects using Aspose.Cells | programmatically clear slicer collections in an .xlsx file with Aspose.Cells C#
// Tags: Aspose.Cells SlicerCollection manipulation | consolidated PDF export from Excel Aspose.Cells | C# worksheet iteration Aspose.Cells | input file existence validation C# | exception handling for Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Correct namespace for SlicerCollection

// Loads an Excel workbook, clears every slicer from each worksheet, and saves the modified workbook as a single PDF, including file existence checks and robust error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and remove all slicers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the slicer collection of the worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Remove slicers starting from the last index to avoid shifting issues
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    slicers.RemoveAt(i);
                }
            }

            // Save the modified workbook as a single consolidated PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

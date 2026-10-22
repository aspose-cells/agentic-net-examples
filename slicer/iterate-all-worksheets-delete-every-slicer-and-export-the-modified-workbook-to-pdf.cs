// Title: Remove all slicers from every worksheet and export the workbook to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates through each worksheet, deletes every slicer in the SlicerCollection, and then exports the workbook to a PDF file. | Show how to safely clear slicer objects from all sheets of a workbook before converting it to PDF with Aspose.Cells in a .NET console application. | Provide a method that verifies the input Excel file exists, removes all slicers across worksheets, and saves the result as a PDF using Aspose.Cells SaveFormat.Pdf.
// Common Searches: asp.net delete slicer objects from all worksheets before PDF conversion using Aspose.Cells | c# code to clear slicer collection in an Excel workbook with Aspose.Cells | how to export an Excel file to PDF after removing slicers in .NET | Aspose.Cells iterate worksheets and clear slicers programmatically | remove Excel slicers programmatically with Aspose.Cells C#
// Tags: Aspose.Cells remove slicers C# | iterate worksheets delete slicer collection | save workbook as PDF Aspose.Cells | C# Excel slicer removal before PDF export | SlicerCollection clear Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// The sample loads an Excel workbook, loops through each worksheet to delete all slicers via the SlicerCollection, and then saves the cleaned workbook as a PDF using Aspose.Cells.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and remove any slicers present
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                SlicerCollection slicers = sheet.Slicers;

                // Remove slicers starting from the end to avoid index shifting
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    slicers.RemoveAt(i);
                }
            }

            // Save the modified workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Rename an Excel defined name from ReportData to SummaryData and save the workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# rename a defined name in an existing workbook | how to change the text of a named range using Aspose.Cells .NET | update Excel named range programmatically and save file with Aspose.Cells | C# code to modify a workbook's name collection and export to a new file | Aspose.Cells example for renaming ReportData to SummaryData
// Tags: defined name modification Aspose.Cells | named range text update C# | Excel name collection manipulation Aspose.Cells | workbook save after name edit Aspose.Cells | Aspose.Cells change defined name property

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel workbook, accesses the defined name "ReportData" from the workbook's name collection, changes its Text property to "SummaryData", and then saves the updated workbook to a new file, with error handling for missing files and exceptions.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "ReportData" from the workbook's name collection
                Name reportName = workbook.Worksheets.Names["ReportData"];
                if (reportName != null)
                {
                    // Change the name of the defined range to "SummaryData"
                    reportName.Text = "SummaryData";
                }
                else
                {
                    Console.WriteLine("Named range 'ReportData' not found.");
                }

                // Save the modified workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

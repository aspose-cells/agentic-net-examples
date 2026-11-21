// Title: Load an Excel workbook, edit a named range, and save as a new file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to change a value in a named range and save as a different workbook | How to access and modify a named range in an existing .xlsx file using Aspose.Cells | Aspose.Cells load workbook, update first cell of MyRange, and export to ReportUpdated.xlsx | Retrieve named range object in C# and set cell value with Aspose.Cells | Save modified Excel file under new name after editing named range with Aspose.Cells .NET
// Tags: Aspose.Cells load workbook C# | modify named range Aspose.Cells | save workbook with new filename Aspose.Cells | named range GetRange method | update cell value Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The C# program verifies the presence of Report.xlsx, loads it with Aspose.Cells, obtains the named range "MyRange", updates the first cell's value to "Updated Value", and saves the result as ReportUpdated.xlsx while handling missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "Report.xlsx";
            const string outputFile = "ReportUpdated.xlsx";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Retrieve the named range (replace "MyRange" with your actual range name)
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange != null)
                {
                    // Get the Aspose.Cells.Range object associated with the name
                    Aspose.Cells.Range range = namedRange.GetRange();

                    // Modify the first cell in the range (adjust as needed)
                    Cell firstCell = range[0, 0];
                    firstCell.PutValue("Updated Value");
                }
                else
                {
                    Console.WriteLine("Named range \"MyRange\" was not found in the workbook.");
                }

                // Save the modified workbook under a new file name
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

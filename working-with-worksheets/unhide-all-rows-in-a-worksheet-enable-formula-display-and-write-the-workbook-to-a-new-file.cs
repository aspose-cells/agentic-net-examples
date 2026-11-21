// Title: C# example: Unhide every row in an Aspose.Cells worksheet and save the workbook as a new file
// Common Searches: C# Aspose.Cells unhide all rows in a worksheet | how to programmatically show hidden rows with Aspose.Cells for .NET | save a modified Excel workbook to a different file using Aspose.Cells C# | iterate over worksheet rows to change visibility with Aspose.Cells | Aspose.Cells .NET alternative to ShowFormula for displaying formulas
// Tags: unhide rows Aspose.Cells C# | worksheet row visibility Aspose.Cells | save workbook to new file Aspose.Cells | iterate rows Aspose.Cells API | display formulas alternative Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The code loads an existing Excel file (or creates a new workbook), accesses the first worksheet, loops through all rows to set IsHidden = false, notes that the ShowFormula property is unavailable, and then saves the updated workbook to a new file while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Load existing workbook or create a new one if the file does not exist.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // default workbook with one worksheet
            }

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Unhide all rows in the worksheet.
            int totalRows = sheet.Cells.Rows.Count;
            for (int i = 0; i < totalRows; i++)
            {
                sheet.Cells.Rows[i].IsHidden = false;
            }

            // NOTE: The ShowFormula property is not available in the current Aspose.Cells version.
            // If needed, alternative approaches can be used to display formulas.

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

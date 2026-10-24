// Title: Hide rows 20‑30 in an Excel worksheet and save the workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide rows 20 through 30 in an Excel file | How to hide a range of rows and save the workbook using Aspose.Cells .NET | C# set Row.IsHidden property for multiple rows with Aspose.Cells | Saving an Excel workbook after concealing rows via Aspose.Cells API | Hide specific rows in Excel programmatically with Aspose.Cells for .NET
// Tags: hide rows Aspose.Cells | Row.IsHidden property C# | Excel hidden rows .NET | save workbook with hidden rows Aspose.Cells | conceal worksheet rows programmatically

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, accesses the first worksheet, iterates over rows 20‑30 (zero‑based indices 19‑29) setting each Row.IsHidden to true, and then saves the file as HiddenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide rows 20 to 30 (Aspose.Cells uses zero‑based indexing)
            for (int rowIndex = 19; rowIndex <= 29; rowIndex++)
            {
                // Use the IsHidden property of the Row object
                sheet.Cells.Rows[rowIndex].IsHidden = true;
            }

            // Define output file path
            string outputPath = "HiddenRows.xlsx";

            // Save the workbook with the hidden rows
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

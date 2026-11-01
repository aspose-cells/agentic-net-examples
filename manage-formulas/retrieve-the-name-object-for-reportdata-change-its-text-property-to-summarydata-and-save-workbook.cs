// Title: Rename a named range from ReportData to SummaryData in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change named range reference text | how to update the RefersTo property of a Name object in Excel with Aspose | rename Excel named range programmatically using Aspose.Cells .NET | set named range to new reference and save workbook Aspose.Cells | C# example for modifying a named range and saving the workbook
// Tags: Aspose.Cells modify named range reference | C# Name.RefersTo property usage | update Excel named range Aspose.Cells | save workbook after named range change | Aspose.Cells .NET Excel named range handling

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, retrieves the named range "ReportData", changes its RefersTo text to "SummaryData", and saves the updated workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Load workbook using the provided load rule
            Workbook workbook = LoadWorkbook("input.xlsx");

            // Retrieve the named range "ReportData"
            Name reportName = workbook.Worksheets.Names["ReportData"];
            if (reportName != null)
            {
                // Change its reference text to "SummaryData"
                reportName.RefersTo = "SummaryData";
            }

            // Save workbook using the provided save rule
            SaveWorkbook(workbook, "output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Load workbook with file existence check
    static Workbook LoadWorkbook(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"The file '{path}' was not found.");

        return new Workbook(path);
    }

    // Save workbook with basic error handling
    static void SaveWorkbook(Workbook wb, string path)
    {
        try
        {
            wb.Save(path);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
            throw;
        }
    }
}

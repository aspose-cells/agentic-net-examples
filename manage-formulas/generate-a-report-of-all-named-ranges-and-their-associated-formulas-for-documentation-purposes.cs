// Title: Create a C# program with Aspose.Cells to generate an Excel report listing all named ranges and their RefersTo formulas
// Common Searches: C# Aspose.Cells list all defined names with their formulas and export to a new workbook | how to generate a named range documentation sheet using Aspose.Cells for .NET | export named ranges and RefersTo expressions from an existing Excel file with Aspose.Cells | Aspose.Cells iterate NameCollection to create a report of named ranges | save named range definitions to a separate Excel file using Aspose.Cells C#
// Tags: Aspose.Cells named range documentation | C# extract named range RefersTo | export NameCollection to new workbook | list defined names Aspose.Cells | generate Excel report of named ranges

using Aspose.Cells;
using System;
using System.IO;

// The example loads a source workbook, iterates through its NameCollection to retrieve each named range's Text and RefersTo formula, writes these details into a new worksheet with headers, and saves the result as a separate Excel file, handling missing files and runtime errors.
class NamedRangeReport
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string reportPath = "NamedRangeReport.xlsx";

            // Verify that the source workbook exists before attempting to load it
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The source file \"{sourcePath}\" was not found.");
                return;
            }

            // Load the source workbook containing the named ranges
            Workbook srcWorkbook = new Workbook(sourcePath);

            // Create a new workbook that will hold the report
            Workbook reportWorkbook = new Workbook();

            // Use the first worksheet for the report and give it a meaningful name
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "NamedRanges";

            // Write column headers
            reportSheet.Cells[0, 0].PutValue("Named Range");
            reportSheet.Cells[0, 1].PutValue("Refers To (Formula)");

            // Retrieve the collection of named ranges from the source workbook
            NameCollection namedRanges = srcWorkbook.Worksheets.Names;

            int rowIndex = 1; // Start writing data from the second row
            foreach (Name namedRange in namedRanges)
            {
                // The name of the defined range (use Text property)
                string rangeName = namedRange.Text;

                // The formula that the name refers to (e.g., =Sheet1!$A$1:$B$5)
                string refersToFormula = namedRange.RefersTo;

                // Populate the report sheet
                reportSheet.Cells[rowIndex, 0].PutValue(rangeName);
                reportSheet.Cells[rowIndex, 1].PutValue(refersToFormula);

                rowIndex++;
            }

            // Save the report workbook to a file
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

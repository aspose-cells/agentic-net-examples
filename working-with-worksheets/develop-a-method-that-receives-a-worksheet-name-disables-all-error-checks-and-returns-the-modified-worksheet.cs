// Title: C# – Disable All Error Checks for a Worksheet Using Aspose.Cells
// Description: A C# helper method that receives a Workbook and a worksheet name, retrieves the matching Worksheet, creates an ErrorCheckOption, disables every ErrorCheckType, applies the option to the entire used range, and returns the updated Worksheet. Useful for removing Excel error indicators before saving or distributing a file.
// Keywords: Aspose.Cells | C# | .NET | disable worksheet error checks | ErrorCheckOption | ErrorCheckType | Excel error indicators | programmatic Excel formatting | workbook manipulation | remove validation warnings
// Common Searches: Aspose.Cells disable all error checks on a sheet | C# turn off Excel error triangles with Aspose.Cells | how to hide error indicators in a worksheet using Aspose.Cells | programmatically disable error checking for a specific worksheet .NET | remove validation warnings from Excel file using Aspose.Cells
// Developer Intent: Create a function that disables every error‑check type on a named worksheet and returns the modified Worksheet.
// Use Cases: Prepare a report workbook for distribution without Excel error triangles. | Generate clean data‑export sheets programmatically, eliminating validation warnings. | Apply consistent error‑check settings across all worksheets in a large workbook.
// AI Prompts: Write a C# method with Aspose.Cells that disables all error checks for a given worksheet name and applies the setting to the whole used range. | Show how to loop through all worksheets in a Workbook and call DisableAllErrorChecks for each sheet. | Explain how to verify that error checks have been disabled after invoking the method in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // A C# helper method that receives a Workbook and a worksheet name, retrieves the matching Worksheet, creates an ErrorCheckOption, disables every ErrorCheckType, applies the option to the entire used range, and returns the updated Worksheet. Useful for removing Excel error indicators before saving or distributing a file.
    public static class WorksheetHelper
    {
        /// <param name="workbook">The workbook containing the worksheet.</param>
        /// <param name="worksheetName">The name of the worksheet to modify.</param>
        /// <returns>The worksheet with all error checks disabled.</returns>
        public static Worksheet DisableAllErrorChecks(Workbook workbook, string worksheetName)
        {
            // Get the worksheet by name; throws if not found.
            Worksheet sheet = workbook.Worksheets[worksheetName];

            // Access the collection of error‑check options for this sheet.
            ErrorCheckOptionCollection options = sheet.ErrorCheckOptions;

            // Add a new ErrorCheckOption to the collection.
            int optionIndex = options.Add();
            ErrorCheckOption option = options[optionIndex];

            // Disable every possible error check type.
            foreach (ErrorCheckType checkType in Enum.GetValues(typeof(ErrorCheckType)))
            {
                option.SetErrorCheck(checkType, false);
            }

            // Apply the option to the whole used range of the worksheet.
            int maxRow = sheet.Cells.MaxRow;
            int maxCol = sheet.Cells.MaxDataColumn; // last column with data
            option.AddRange(CellArea.CreateCellArea(0, 0, maxRow, maxCol));

            return sheet;
        }
    }

    public static class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                // Disable all error checks on the first worksheet.
                string sheetName = workbook.Worksheets[0].Name;
                WorksheetHelper.DisableAllErrorChecks(workbook, sheetName);

                // Save the modified workbook.
                string outputPath = "output.xlsx";
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

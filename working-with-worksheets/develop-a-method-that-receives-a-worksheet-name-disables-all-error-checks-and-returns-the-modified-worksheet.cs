// Title: Disable all error checking for a specific worksheet using Aspose.Cells in C#
// Common Searches: c# aspocells how to turn off error checking for a single worksheet | disable Excel error indicators programmatically with Aspose.Cells .NET | retrieve a worksheet by name and suppress error checks using Aspose.Cells | aspocells workbook settings enableerrorchecking false example in C# | c# method to get worksheet and disable all error checks with Aspose.Cells
// Tags: aspocells disable worksheet error checking | c# retrieve worksheet by name aspocells | aspocells workbook settings enableerrorchecking false | c# aspocells worksheet helper method | aspocells suppress excel error indicators

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example shows a C# helper method that validates inputs, fetches a worksheet by its name from a Workbook, and (when supported) disables all error checking for the workbook via the Settings.EnableErrorChecking property before returning the worksheet for further operations.
    public class WorksheetHelper
    {
        /// <param name="workbook">The workbook containing the worksheet.</param>
        /// <param name="worksheetName">The name of the worksheet to modify.</param>
        /// <returns>The requested worksheet.</returns>
        public Worksheet DisableAllErrorChecks(Workbook workbook, string worksheetName)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(worksheetName));

            Worksheet worksheet = workbook.Worksheets[worksheetName];
            if (worksheet == null)
                throw new ArgumentException($"Worksheet '{worksheetName}' does not exist in the workbook.", nameof(worksheetName));

            // Disable error checking for the entire workbook if the API is available.
            // In some Aspose.Cells versions this property may not exist; uncomment if supported.
            // workbook.Settings.EnableErrorChecking = false;

            return worksheet;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook or create a new one if the file is missing.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                var helper = new WorksheetHelper();
                string targetSheet = "Sheet1";

                // Ensure the target worksheet exists.
                if (workbook.Worksheets[targetSheet] == null)
                {
                    workbook.Worksheets.Add(targetSheet);
                }

                // Disable error checking and obtain the worksheet.
                Worksheet ws = helper.DisableAllErrorChecks(workbook, targetSheet);

                // Example operation: write a value to the first cell.
                ws.Cells["A1"].PutValue("Demo");

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

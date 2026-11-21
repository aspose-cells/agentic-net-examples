// Title: How to lock an Excel slicer on a protected worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock slicer after protecting worksheet | prevent end users from changing slicer settings in Excel with Aspose.Cells | set slicer IsLocked property in Aspose.Cells .NET example | protect sheet and lock slicer using Aspose.Cells API | C# code to lock slicer on protected Excel sheet with Aspose.Cells
// Tags: Aspose.Cells slicer IsLocked usage | C# protect worksheet while locking slicer | Excel slicer lock with Aspose.Cells API | C# lock slicer on a protected worksheet | Aspose.Cells worksheet protection with slicer lock

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for Slicer class

namespace AsposeCellsExample
{
    // The example loads an existing workbook, locates a specific slicer, sets its IsLocked property to true, protects the worksheet with a password using ProtectionType.All, and saves the result, ensuring end users cannot modify the slicer's configuration.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string sheetName = "Sheet1";          // adjust if needed
            const string slicerName = "Slicer_Name";    // replace with actual slicer name
            const string password = "password123";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the worksheet that contains the slicer
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                    return;
                }

                // Retrieve the slicer by its name
                Slicer slicer = sheet.Slicers[slicerName];
                if (slicer == null)
                {
                    Console.WriteLine($"Slicer \"{slicerName}\" not found on sheet \"{sheetName}\".");
                    return;
                }

                // Lock the slicer so users cannot change its configuration
                slicer.IsLocked = true;

                // Protect the worksheet (specify the protection type, password, and old password)
                sheet.Protect(ProtectionType.All, password, string.Empty);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

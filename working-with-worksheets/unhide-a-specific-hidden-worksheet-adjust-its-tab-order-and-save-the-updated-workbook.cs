// Title: Unhide and Reorder a Worksheet with Aspose.Cells for .NET (C#)
// Description: Loads an existing XLSX file, verifies its presence, makes a specified hidden worksheet visible, moves it to a given zero‑based tab index, and saves the modified workbook to a new file using Aspose.Cells.
// Keywords: Aspose.Cells | C# | unhide worksheet | move worksheet tab | reorder Excel sheets | set worksheet visibility | Workbook.Save | Excel automation | programmatic sheet order | hidden sheet handling
// Common Searches: Aspose.Cells unhide hidden sheet C# | change worksheet tab order with Aspose.Cells | move sheet after making it visible Aspose.Cells .NET | how to make a hidden worksheet visible programmatically | reorder Excel worksheets using Aspose.Cells
// Developer Intent: Make a hidden worksheet visible, change its position in the tab order, and persist the changes.
// Use Cases: Expose a confidential sheet before distributing a report. | Adjust the sequence of generated sheets to match business requirements. | Correct tab ordering after dynamically adding or removing worksheets. | Safely handle missing input files or absent target worksheets with clear logging. | Validate the target index against the workbook's sheet count to prevent runtime errors.
// AI Prompts: Write C# code using Aspose.Cells that checks for an existing workbook, unhides a sheet named "Report", moves it to the first tab, and saves the file as "updated.xlsx". | Provide a robust Aspose.Cells example that validates the input path, makes a hidden worksheet visible, repositions it to a specified index, and handles invalid index or missing sheet scenarios. | Create a reusable method in C# that accepts a file path, sheet name, and target index, then uses Aspose.Cells to set the sheet's visibility, reorder it, and return the path of the saved workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing XLSX file, verifies its presence, makes a specified hidden worksheet visible, moves it to a given zero‑based tab index, and saves the modified workbook to a new file using Aspose.Cells.
    public class UnhideAndReorderWorksheet
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the existing workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the worksheet to unhide
            string sheetNameToUnhide = "HiddenSheet";

            // Desired new position (0‑based index) in the tab order
            int newTabIndex = 1;

            // Find the worksheet by name
            Worksheet sheet = workbook.Worksheets[sheetNameToUnhide];

            if (sheet != null)
            {
                // Unhide the worksheet
                sheet.IsVisible = true; // or: sheet.SetVisible(true, true);

                // Move the worksheet to the desired tab position if the index is valid
                if (newTabIndex >= 0 && newTabIndex < workbook.Worksheets.Count)
                {
                    sheet.MoveTo(newTabIndex);
                }
                else
                {
                    Console.WriteLine("Invalid tab index specified.");
                }
            }
            else
            {
                Console.WriteLine($"Worksheet \"{sheetNameToUnhide}\" not found.");
            }

            // Save the updated workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with the worksheet unhidden and reordered.");
        }
    }
}

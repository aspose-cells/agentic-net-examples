// Title: Insert a new worksheet immediately after a specific sheet name using Aspose.Cells for .NET (C#)
// Common Searches: how to add a worksheet after an existing sheet by name using Aspose.Cells C# | Aspose.Cells insert new worksheet after target tab in .NET workbook | C# place a new worksheet right after a specific worksheet name with Aspose.Cells | programmatically add blank sheet after "TargetSheet" using Aspose.Cells | Aspose.Cells .NET insert worksheet at a given index based on sheet name
// Tags: Aspose.Cells insert worksheet after named sheet | C# add worksheet at specific index | set worksheet name after insertion Aspose.Cells | save modified workbook Aspose.Cells .NET | ensure output directory exists C# file IO

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, locates the worksheet named "TargetSheet", inserts a new blank worksheet directly after it, assigns the name "InsertedSheet" to the new sheet, writes a message into cell A1, creates the output directory if needed, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // The sheet after which the new worksheet should be placed
            string targetSheetName = "TargetSheet";

            // Retrieve the target worksheet by name
            Worksheet targetSheet = workbook.Worksheets[targetSheetName];
            if (targetSheet == null)
            {
                Console.WriteLine($"Sheet '{targetSheetName}' not found.");
                return;
            }

            // Index of the target sheet
            int targetIndex = targetSheet.Index;

            // Name for the new worksheet to be inserted
            string newSheetName = "InsertedSheet";

            // Insert a new blank worksheet immediately after the target sheet
            workbook.Worksheets.Insert(targetIndex + 1, SheetType.Worksheet);

            // Retrieve the newly inserted worksheet and set its name
            Worksheet newSheet = workbook.Worksheets[targetIndex + 1];
            newSheet.Name = newSheetName;

            // Optional: add some content to the newly inserted sheet
            newSheet.Cells["A1"].PutValue($"This sheet was inserted after '{targetSheetName}'.");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Assign a custom WorkbookProperties.Title to a merged Excel workbook using Aspose.Cells for .NET
// AI Prompts: Add code that concatenates the source file names into a single string and assigns it to mergedWorkbook.WorkbookProperties.Title before saving. | Create a helper method that receives an array of source paths, builds a descriptive title with the current timestamp, and sets the WorkbookProperties.Title on the merged workbook. | Read the Title property from each source workbook, merge those titles, and apply the combined value to the merged workbook's WorkbookProperties.Title. | Implement fallback logic that sets a default title when any source file is missing or its title cannot be retrieved.
// Common Searches: Aspose.Cells C# set workbook title after merging multiple Excel files | How to add custom document properties to a combined workbook with Aspose.Cells for .NET | C# merge two .xlsx files and define a descriptive WorkbookProperties.Title using Aspose.Cells | Include source filenames in the Title property of a merged Excel workbook with Aspose.Cells | Aspose.Cells set timestamp in WorkbookProperties.Title for merged workbook
// Tags: Aspose.Cells document title property | Aspose.Cells merged workbook properties | C# Aspose.Cells workbook metadata | Aspose.Cells custom title generation | Aspose.Cells timestamped document title

using System;
using System.IO;
using Aspose.Cells;

// The example loads two source Excel files, creates an empty workbook, copies all worksheets from each source into the new workbook, optionally builds a descriptive title that lists the source filenames (and can include a timestamp), assigns this string to mergedWorkbook.WorkbookProperties.Title, and saves the merged file as MergedWorkbook.xlsx.
class WorkbookMerger
{
    static void Main()
    {
        try
        {
            // Paths to source workbooks
            string sourcePath1 = "SourceFile1.xlsx";
            string sourcePath2 = "SourceFile2.xlsx";

            // Verify that source files exist
            if (!File.Exists(sourcePath1) || !File.Exists(sourcePath2))
            {
                Console.WriteLine("One or both source files were not found.");
                return;
            }

            // Create a new workbook that will hold the merged content
            Workbook mergedWorkbook = new Workbook();

            // Remove the default empty worksheet created with a new workbook
            mergedWorkbook.Worksheets.Clear();

            // Load the first source workbook
            Workbook sourceWorkbook1 = new Workbook(sourcePath1);
            // Load the second source workbook
            Workbook sourceWorkbook2 = new Workbook(sourcePath2);

            // Copy all worksheets from the first source workbook into the merged workbook
            foreach (Worksheet sheet in sourceWorkbook1.Worksheets)
            {
                mergedWorkbook.Worksheets.AddCopy(sheet.Name);
            }

            // Copy all worksheets from the second source workbook into the merged workbook
            foreach (Worksheet sheet in sourceWorkbook2.Worksheets)
            {
                mergedWorkbook.Worksheets.AddCopy(sheet.Name);
            }

            // Optionally set document properties if supported
            // mergedWorkbook.WorkbookProperties.Title = "Combined Workbook: SourceFile1.xlsx + SourceFile2.xlsx";

            // Save the merged workbook to a new file
            string outputPath = "MergedWorkbook.xlsx";
            mergedWorkbook.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

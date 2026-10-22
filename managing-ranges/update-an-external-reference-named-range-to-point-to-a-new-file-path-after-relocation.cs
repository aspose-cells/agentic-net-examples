// Title: How to update the file path of an external named range in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change external named range file path after moving source workbook | Update RefersTo of an external named range to a new location with Aspose.Cells .NET | Replace old workbook path in a named range reference using Aspose.Cells in C# | Modify external reference of a workbook‑level named range in Aspose.Cells for .NET | C# code to relocate external named range source file in Excel using Aspose.Cells
// Tags: update external named range RefersTo Aspose.Cells | replace workbook path in Excel named range C# | modify external reference of workbook‑level named range | Aspose.Cells external link path change | C# regex extract file path from RefersTo

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads a workbook, finds a workbook‑level named range that points to an external workbook, extracts the old file path from its RefersTo string with a regular expression, replaces it with a new path, updates the RefersTo property, and saves the workbook to a new file.
class UpdateExternalReference
{
    static void Main()
    {
        try
        {
            // Paths
            string inputPath = @"C:\Workbooks\MainWorkbook.xlsx";
            string outputPath = @"C:\Workbooks\MainWorkbook_Updated.xlsx";
            string newFilePath = @"D:\NewFolder\Source.xlsx";
            string externalRangeName = "MyExternalRange";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of named ranges at the workbook level
            Name externalName = workbook.Worksheets.Names[externalRangeName];

            if (externalName != null)
            {
                // Example RefersTo: ='[C:\OldFolder\Source.xlsx]Sheet1'!$A$1
                string currentRef = externalName.RefersTo;

                // Extract the file path inside brackets [...]
                Match match = Regex.Match(currentRef, @"\[(.*?)\]");
                if (match.Success)
                {
                    string oldFilePath = match.Groups[1].Value;
                    string updatedRef = currentRef.Replace(oldFilePath, newFilePath);
                    externalName.RefersTo = updatedRef;
                }
                else
                {
                    Console.WriteLine("The named range does not contain an external reference.");
                }
            }
            else
            {
                Console.WriteLine($"Named range '{externalRangeName}' not found.");
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

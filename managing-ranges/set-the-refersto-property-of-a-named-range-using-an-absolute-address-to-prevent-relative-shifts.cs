// Title: Set the RefersTo property of a named range to an absolute address in Aspose.Cells for .NET
// AI Prompts: Create a Name called MyAbsoluteRange and assign its RefersTo property the absolute range $A$1:$B$10 using Aspose.Cells. | Add a named range to the worksheet's Names collection and lock it with an absolute address so it does not move when rows or columns are inserted. | Ensure the output folder exists, then save the workbook and verify that the named range retains the absolute reference.
// Common Searches: Aspose.Cells how to set a named range with an absolute address | prevent named range from shifting after inserting rows in .NET Excel library | RefersTo property absolute reference example for Aspose.Cells | save workbook with fixed named range using Aspose.Cells for C#
// Tags: Aspose.Cells RefersTo absolute address | named range fixed reference .NET | add named range to worksheet Names collection | save workbook with named range Aspose.Cells | prevent range shift on row insert Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, adds sample data, defines an absolute address "$A$1:$B$10", adds a named range "MyAbsoluteRange" to the worksheet's Names collection, sets its RefersTo property to the absolute address, ensures the output directory exists, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Optional: put some sample data
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["B10"].PutValue(10);

            // Define the absolute address for the named range
            // The $ signs make the address absolute, preventing relative shifts
            string absoluteAddress = "$A$1:$B$10";

            // Add a named range to the workbook's Names collection
            // Add returns the index of the newly added name
            int nameIndex = workbook.Worksheets.Names.Add("MyAbsoluteRange");
            Name namedRange = workbook.Worksheets.Names[nameIndex];

            // Set the RefersTo property using the absolute address
            namedRange.RefersTo = absoluteAddress;

            // Determine output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists (prevents FileNotFoundException on save)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

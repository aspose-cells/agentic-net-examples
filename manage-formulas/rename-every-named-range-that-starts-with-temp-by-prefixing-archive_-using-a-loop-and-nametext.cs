// Title: Rename Excel named ranges that start with 'Temp' to 'Archive_' using Aspose.Cells for .NET (C#)
// Common Searches: C# code to rename Excel defined names that begin with Temp using Aspose.Cells | How to add a prefix to specific named ranges in a workbook with Aspose.Cells | Loop through workbook Names collection and change Name.Text in Aspose.Cells C# | Batch rename temporary named ranges to archive format in .NET Excel processing | Rename named ranges with a certain prefix in an existing XLSX file using Aspose.Cells
// Tags: rename defined names Aspose.Cells C# | prefix named ranges Excel Aspose.Cells | iterate workbook names collection .NET | modify Name.Text property Aspose.Cells | batch rename temporary ranges XLSX

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, iterates over all defined names, and for each name that starts with "Temp" it replaces the prefix with "Archive_" by setting the Name.Text property, then saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        // Ensure the input file exists; otherwise create a new workbook.
        if (File.Exists(inputPath))
        {
            try
            {
                // Try loading the workbook normally.
                workbook = new Workbook(inputPath);
            }
            catch (Exception)
            {
                try
                {
                    // Attempt to load with an empty password (for password‑protected files).
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = "" };
                    workbook = new Workbook(inputPath, loadOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading workbook: {ex.Message}");
                    return;
                }
            }
        }
        else
        {
            // Create a new workbook with a default worksheet.
            workbook = new Workbook();
            workbook.Worksheets[0].Name = "Sheet1";
        }

        try
        {
            // Iterate over all defined names in the workbook.
            for (int i = 0; i < workbook.Worksheets.Names.Count; i++)
            {
                Name definedName = workbook.Worksheets.Names[i];

                // Rename names that start with "Temp" by prefixing "Archive_".
                if (definedName.Text.StartsWith("Temp", StringComparison.Ordinal))
                {
                    string newName = "Archive_" + definedName.Text.Substring(4);
                    definedName.Text = newName;
                }
            }

            // Save the updated workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}

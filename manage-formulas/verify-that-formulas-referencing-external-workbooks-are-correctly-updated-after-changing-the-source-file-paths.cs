// Title: Updating external workbook links after moving the source file and verifying formula results with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# update external workbook link after moving source file | verify external reference formula value after changing source path using Aspose.Cells | recalculate Excel formulas with new external file location in C# | how to modify external reference to full file path in Aspose.Cells workbook | check external link update success programmatically with Aspose.Cells .NET
// Tags: update external workbook link Aspose.Cells | recalculate formulas after external reference change C# | set full file path in external Excel reference Aspose.Cells | validate external reference result .NET | adjust external link file location programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example creates a source workbook, adds a value, saves a destination workbook with an external reference to the source, moves the source file to a new folder, updates the destination workbook's formula to use the full new path, recalculates the formula, and verifies that the result matches the original value before cleaning up temporary files.
class ExternalReferenceVerification
{
    static void Main()
    {
        // Paths for the source and destination workbooks
        string sourcePath = "source.xlsx";
        string destPath = "dest.xlsx";
        string newFolder = "NewSourceFolder";
        string newSourcePath = Path.Combine(newFolder, "source.xlsx");

        try
        {
            // -------------------------------------------------
            // Create source workbook with a value in A1
            // -------------------------------------------------
            Workbook sourceWb = new Workbook(); // create
            Worksheet sourceSheet = sourceWb.Worksheets[0];
            sourceSheet.Cells["A1"].PutValue(123.45); // sample value
            sourceWb.Save(sourcePath); // save

            // -------------------------------------------------
            // Create destination workbook with a formula referencing the source workbook
            // -------------------------------------------------
            Workbook destWb = new Workbook(); // create
            Worksheet destSheet = destWb.Worksheets[0];
            // External reference formula: ='[source.xlsx]Sheet1'!A1
            destSheet.Cells["A1"].Formula = $"='[{Path.GetFileName(sourcePath)}]Sheet1'!A1";
            destWb.Save(destPath); // save

            // -------------------------------------------------
            // Move the source workbook to a new folder (simulating a path change)
            // -------------------------------------------------
            Directory.CreateDirectory(newFolder);
            if (File.Exists(sourcePath))
            {
                File.Move(sourcePath, newSourcePath);
            }

            // -------------------------------------------------
            // Load the destination workbook and update external links to the new source path
            // -------------------------------------------------
            if (!File.Exists(destPath))
                throw new FileNotFoundException($"Destination workbook not found: {destPath}");

            Workbook loadedDestWb = new Workbook(destPath); // load

            // Update the external reference formula to point to the new location
            Worksheet loadedDestSheet = loadedDestWb.Worksheets[0];
            // Construct a formula with full path to the moved source workbook
            // Excel external reference format: ='C:\Path\Folder\[FileName.xlsx]Sheet1'!A1
            string fullPathFormula = $"='[{newSourcePath}]Sheet1'!A1";
            loadedDestSheet.Cells["A1"].Formula = fullPathFormula;

            // Recalculate formulas after updating the reference
            loadedDestWb.CalculateFormula();

            // -------------------------------------------------
            // Verify that the formula result reflects the value from the moved source workbook
            // -------------------------------------------------
            double result = loadedDestWb.Worksheets[0].Cells["A1"].DoubleValue;
            Console.WriteLine($"Formula result after updating external link: {result}");

            // Expected result should be 123.45
            if (Math.Abs(result - 123.45) < 0.0001)
            {
                Console.WriteLine("Verification succeeded: external reference updated correctly.");
            }
            else
            {
                Console.WriteLine("Verification failed: external reference not updated as expected.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up (optional)
            try
            {
                if (File.Exists(newSourcePath))
                    File.Delete(newSourcePath);
                if (File.Exists(destPath))
                    File.Delete(destPath);
                if (Directory.Exists(newFolder))
                    Directory.Delete(newFolder);
            }
            catch
            {
                // Suppress any cleanup errors
            }
        }
    }
}

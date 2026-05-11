using System;
using Aspose.Cells;

class MergeWorkbooks
{
    static void Main()
    {
        // Define the paths of the workbooks that need to be merged.
        string[] sourceFiles = { "File1.xlsx", "File2.xlsx", "File3.xlsx" };

        // Load the first workbook; it will serve as the destination workbook.
        // Uses the Workbook(string) constructor as defined in the rules.
        Workbook destWorkbook = new Workbook(sourceFiles[0]);

        // Iterate over the remaining source files, load each, and combine it into the destination.
        for (int i = 1; i < sourceFiles.Length; i++)
        {
            // Load a source workbook using the same constructor rule.
            Workbook srcWorkbook = new Workbook(sourceFiles[i]);

            // Combine the source workbook into the destination workbook.
            // This utilizes the Workbook.Combine(Workbook) method rule.
            destWorkbook.Combine(srcWorkbook);
        }

        // Save the merged workbook to a new file.
        // The Save(string, SaveFormat) method follows the provided save rule.
        destWorkbook.Save("MergedOutput.xlsx", SaveFormat.Xlsx);
    }
}
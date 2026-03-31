using System;
using Aspose.Cells;

class CombineWorkbooksDemo
{
    static void Main()
    {
        // List of workbook files to be merged
        string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

        // Create an empty destination workbook
        Workbook destWorkbook = new Workbook();

        // Load each source workbook and combine it into the destination workbook
        foreach (string filePath in sourceFiles)
        {
            // Load the source workbook from file
            Workbook srcWorkbook = new Workbook(filePath);

            // Merge the source workbook into the destination workbook
            destWorkbook.Combine(srcWorkbook);

            // Release resources used by the source workbook
            srcWorkbook.Dispose();
        }

        // Save the combined workbook to disk
        destWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources used by the destination workbook
        destWorkbook.Dispose();
    }
}
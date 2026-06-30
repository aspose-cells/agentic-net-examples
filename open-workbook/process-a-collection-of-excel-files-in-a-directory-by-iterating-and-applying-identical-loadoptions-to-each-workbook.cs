using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – processes all Excel files in a folder using the same LoadOptions.
class Program
{
    static void Main()
    {
        // Folder containing the Excel files to process
        string folderPath = @"C:\ExcelFiles";

        // Create a single LoadOptions instance that will be reused for every workbook
        LoadOptions loadOptions = new LoadOptions
        {
            // Example setting – ignore data that is not marked for printing
            IgnoreNotPrinted = true
        };

        // Iterate over all .xlsx files in the directory (add other extensions if needed)
        foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
        {
            // Load the workbook with the predefined LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example processing: display the number of worksheets in the console
            Console.WriteLine($"{Path.GetFileName(filePath)} – Worksheets: {workbook.Worksheets.Count}");

            // If modifications are made, you could save the workbook like this:
            // string outputPath = Path.Combine(folderPath, "Processed_" + Path.GetFileName(filePath));
            // workbook.Save(outputPath);
        }
    }
}
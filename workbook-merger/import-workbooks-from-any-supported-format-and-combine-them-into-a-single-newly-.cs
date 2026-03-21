using System;
using System.IO;
using Aspose.Cells;

public class CombineWorkbooksDemo
{
    public static void Run(string[] sourceFiles, string outputFile)
    {
        // Create an empty destination workbook.
        Workbook destWorkbook = new Workbook();

        foreach (string filePath in sourceFiles)
        {
            if (!File.Exists(filePath))
                continue; // Skip missing files.

            // Determine appropriate load options based on file extension.
            LoadOptions loadOptions;
            string ext = Path.GetExtension(filePath).ToLowerInvariant();

            if (ext == ".csv")
                loadOptions = new LoadOptions(LoadFormat.Csv);
            else
                loadOptions = new LoadOptions(LoadFormat.Auto);

            // Load the source workbook.
            Workbook srcWorkbook = new Workbook(filePath, loadOptions);

            // Merge the source workbook into the destination workbook.
            destWorkbook.Combine(srcWorkbook);
        }

        // Save the combined workbook.
        destWorkbook.Save(outputFile, SaveFormat.Xlsx);
    }

    public static void Main()
    {
        string[] filesToCombine = new string[]
        {
            "Report1.xlsx",
            "Report2.xls",
            "Data.csv"
        };

        string combinedFile = "CombinedReport.xlsx";

        Run(filesToCombine, combinedFile);

        Console.WriteLine($"Workbooks combined successfully into '{combinedFile}'.");
    }
}
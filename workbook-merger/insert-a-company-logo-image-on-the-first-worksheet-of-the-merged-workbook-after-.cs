using System;
using System.IO;
using Aspose.Cells;

class MergeAndAddLogo
{
    static void Main()
    {
        // Files to be merged
        string[] sourceFiles = { "File1.xlsx", "File2.xlsx" };

        // Path to the company logo image
        string logoPath = "logo.png";

        // Output file name
        string outputFile = "MergedWithLogo.xlsx";

        // Create an empty workbook that will hold the merged content
        Workbook mergedWorkbook = new Workbook();

        // Combine each source workbook into the merged workbook
        foreach (string file in sourceFiles)
        {
            Workbook srcWorkbook = new Workbook(file);
            mergedWorkbook.Combine(srcWorkbook);
        }

        // Insert the logo image into the first worksheet if the file exists
        Worksheet firstSheet = mergedWorkbook.Worksheets[0];
        if (File.Exists(logoPath))
        {
            using (FileStream fs = new FileStream(logoPath, FileMode.Open, FileAccess.Read))
            {
                // The picture will occupy cells B2 (row 1, column 1) to D6 (row 5, column 5).
                firstSheet.Pictures.Add(1, 1, 5, 5, fs);
            }
        }

        // Save the merged workbook with the logo (if added)
        mergedWorkbook.Save(outputFile, SaveFormat.Xlsx);
    }
}
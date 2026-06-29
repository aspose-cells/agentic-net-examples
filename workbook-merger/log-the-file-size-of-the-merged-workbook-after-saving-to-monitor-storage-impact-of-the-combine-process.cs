using System;
using System.IO;
using Aspose.Cells;

class MergeAndLogSize
{
    static void Main()
    {
        // Create the first workbook and add sample data
        using (Workbook wb1 = new Workbook())
        {
            wb1.Worksheets[0].Cells["A1"].PutValue("Data from first workbook");

            // Create the second workbook and add sample data
            using (Workbook wb2 = new Workbook())
            {
                wb2.Worksheets[0].Cells["A1"].PutValue("Data from second workbook");

                // Combine the second workbook into the first workbook
                wb1.Combine(wb2);

                // Define the output file path
                string outputPath = "CombinedOutput.xlsx";

                // Save the combined workbook (uses the provided Save method)
                wb1.Save(outputPath, SaveFormat.Xlsx);

                // Log the file size after saving
                FileInfo fileInfo = new FileInfo(outputPath);
                Console.WriteLine($"Combined workbook saved to '{outputPath}'. Size: {fileInfo.Length} bytes.");
            }
        }
    }
}
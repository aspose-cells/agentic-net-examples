using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class AddProcessedDateProperty
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Add a custom document property named "ProcessedDate" with the current date and time
        DateTime now = DateTime.Now;
        workbook.CustomDocumentProperties.Add("ProcessedDate", now);

        // Save the workbook with the new property (replace with desired output path)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
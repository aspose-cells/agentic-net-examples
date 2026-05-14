using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class AddProcessedDateProperty
{
    static void Main()
    {
        // Path to the existing workbook to open
        string inputPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Add a custom document property named "ProcessedDate" with the current date and time
        workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

        // Save the modified workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
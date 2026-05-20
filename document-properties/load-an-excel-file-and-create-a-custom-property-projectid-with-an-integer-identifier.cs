using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing Excel file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Add a custom document property named "ProjectId" with an integer value
        int projectId = 12345; // example identifier
        workbook.CustomDocumentProperties.Add("ProjectId", projectId);

        // Save the workbook with the new property
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
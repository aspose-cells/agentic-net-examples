using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class RemoveCustomProperty
{
    static void Main()
    {
        // Paths for the source and the cleaned workbook
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook from a file stream
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(stream);

            // If the custom property "IsReviewed" exists, remove it
            if (workbook.CustomDocumentProperties.Contains("IsReviewed"))
            {
                workbook.CustomDocumentProperties.Remove("IsReviewed");
            }

            // Save the workbook with the property removed
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
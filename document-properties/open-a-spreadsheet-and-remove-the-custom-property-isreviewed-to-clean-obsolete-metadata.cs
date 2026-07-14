using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class RemoveCustomProperty
{
    static void Main()
    {
        // Paths for the source and destination files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Open the workbook from a file stream
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(stream);

            // Remove the custom property named "IsReviewed" if it exists
            if (workbook.CustomDocumentProperties.Contains("IsReviewed"))
            {
                workbook.CustomDocumentProperties.Remove("IsReviewed");
            }

            // Save the updated workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
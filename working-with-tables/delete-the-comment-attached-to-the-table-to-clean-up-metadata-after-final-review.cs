using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace DeleteTableCommentApp
{
    class DeleteTableComment
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through worksheets and their tables (ListObjects)
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    foreach (ListObject table in worksheet.ListObjects)
                    {
                        // Remove any comment attached to the table
                        table.Comment = string.Empty;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
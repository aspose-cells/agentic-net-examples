using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveCustomPropertyDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook from file stream
                using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                {
                    Workbook workbook = new Workbook(stream);

                    // Remove custom property if it exists
                    if (workbook.CustomDocumentProperties.Contains("IsReviewed"))
                    {
                        workbook.CustomDocumentProperties.Remove("IsReviewed");
                    }

                    // Save the modified workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                }

                Console.WriteLine($"Custom property \"IsReviewed\" removed and file saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
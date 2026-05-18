using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    public class WorkbookCombineExample
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Destination workbook that will receive merged content
            using (Workbook destWorkbook = new Workbook())
            {
                // First source workbook with sample data
                using (Workbook source1 = new Workbook())
                {
                    Worksheet sheet1 = source1.Worksheets[0];
                    sheet1.Name = "Source1";
                    sheet1.Cells["A1"].PutValue("Data from source workbook 1");
                    sheet1.Cells["B2"].PutValue(123);

                    // Merge first source into destination
                    destWorkbook.Combine(source1);
                }

                // Second source workbook with sample data
                using (Workbook source2 = new Workbook())
                {
                    Worksheet sheet2 = source2.Worksheets[0];
                    sheet2.Name = "Source2";
                    sheet2.Cells["A1"].PutValue("Data from source workbook 2");
                    sheet2.Cells["C3"].PutValue(DateTime.Now);

                    // Merge second source into destination
                    destWorkbook.Combine(source2);
                }

                // Define output path
                string outputPath = "CombinedWorkbook.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the combined workbook
                destWorkbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbooks combined and saved as '{outputPath}'.");
            }
        }
    }
}
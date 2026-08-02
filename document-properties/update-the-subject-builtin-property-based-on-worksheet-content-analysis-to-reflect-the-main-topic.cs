using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class UpdateSubjectBasedOnContentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Quarterly Sales Report");
                sheet.Cells["A2"].PutValue("Region");
                sheet.Cells["B2"].PutValue("Sales");
                sheet.Cells["A3"].PutValue("North");
                sheet.Cells["B3"].PutValue(120000);
                sheet.Cells["A4"].PutValue("South");
                sheet.Cells["B4"].PutValue(95000);

                // Determine main topic: first non‑empty cell in column A
                string mainTopic = null;
                int lastRow = sheet.Cells.MaxDataRow;
                for (int row = 0; row <= lastRow; row++)
                {
                    var cell = sheet.Cells[row, 0]; // Column A (index 0)
                    if (cell != null && cell.Type != CellValueType.IsNull && !string.IsNullOrWhiteSpace(cell.StringValue))
                    {
                        mainTopic = cell.StringValue;
                        break;
                    }
                }

                // Set the Subject built‑in property if a topic was found
                if (!string.IsNullOrEmpty(mainTopic))
                {
                    workbook.BuiltInDocumentProperties.Subject = mainTopic;
                }

                // Save the workbook
                string outputPath = "SubjectUpdatedWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Verify the Subject property by loading the saved file
                if (File.Exists(outputPath))
                {
                    Workbook loaded = new Workbook(outputPath);
                    Console.WriteLine("Subject property set to: " + loaded.BuiltInDocumentProperties.Subject);
                }
                else
                {
                    Console.WriteLine($"Error: The file '{outputPath}' was not found after saving.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                UpdateSubjectBasedOnContentDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }
    }
}
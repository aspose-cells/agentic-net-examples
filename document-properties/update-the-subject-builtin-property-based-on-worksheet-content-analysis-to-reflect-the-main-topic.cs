using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsSubjectUpdateDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data – the first non‑empty cell will be used as the main topic
            sheet.Cells["A1"].PutValue("Quarterly Sales Analysis"); // This will become the Subject
            sheet.Cells["A2"].PutValue("Region");
            sheet.Cells["B2"].PutValue("Revenue");
            sheet.Cells["A3"].PutValue("North");
            sheet.Cells["B3"].PutValue(125000);
            sheet.Cells["A4"].PutValue("South");
            sheet.Cells["B4"].PutValue(98000);

            // Determine the main topic by scanning the first row for the first non‑empty string value
            string mainTopic = "Untitled Document"; // fallback value
            Cells cells = sheet.Cells;
            int maxColumn = cells.MaxColumn; // number of used columns in the sheet

            for (int col = 0; col <= maxColumn; col++)
            {
                object value = cells[0, col].Value; // first row (index 0)
                if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                {
                    mainTopic = value.ToString();
                    break;
                }
            }

            // Update the built‑in Subject property with the derived main topic
            workbook.BuiltInDocumentProperties.Subject = mainTopic;

            // Optionally, also set the Title to the same value for consistency
            workbook.BuiltInDocumentProperties.Title = mainTopic;

            // Save the workbook
            string outputPath = "SubjectUpdatedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Output the result to the console for verification
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine($"Subject property set to: {workbook.BuiltInDocumentProperties.Subject}");
        }
    }
}
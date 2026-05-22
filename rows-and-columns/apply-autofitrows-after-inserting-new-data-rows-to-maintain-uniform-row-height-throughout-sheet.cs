using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AutoFitRowsAfterInsertDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Add initial data
                cells["A1"].PutValue("Header");
                cells["B1"].PutValue("Description");
                cells["A2"].PutValue("Item 1");
                cells["B2"].PutValue("This is a longer description that may wrap and increase row height.");
                cells["A3"].PutValue("Item 2");
                cells["B3"].PutValue("Short desc");

                // Insert two new rows at index 2 (between rows 2 and 3)
                cells.InsertRows(2, 2);

                // Populate the newly inserted rows
                cells["A3"].PutValue("Inserted Item A");
                cells["B3"].PutValue("Inserted description with enough text to wrap onto multiple lines.");
                cells["A4"].PutValue("Inserted Item B");
                cells["B4"].PutValue("Another inserted description.");

                // Auto‑fit all rows so their heights match the content
                worksheet.AutoFitRows();

                // Ensure uniform row height by applying the maximum height found
                double maxHeight = 0;
                int lastRow = worksheet.Cells.MaxDataRow;
                for (int i = 0; i <= lastRow; i++)
                {
                    double h = worksheet.Cells.GetRowHeight(i);
                    if (h > maxHeight) maxHeight = h;
                }
                for (int i = 0; i <= lastRow; i++)
                {
                    worksheet.Cells.Rows[i].Height = maxHeight;
                }

                // Save the workbook
                string outputPath = "AutoFitRowsAfterInsertDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            AutoFitRowsAfterInsertDemo.Run();
        }
    }
}
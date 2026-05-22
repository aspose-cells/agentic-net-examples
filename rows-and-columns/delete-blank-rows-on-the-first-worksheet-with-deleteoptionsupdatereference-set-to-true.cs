using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DeleteBlankRowsWithUpdateReference
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Add sample data with blank rows (rows 3 and 5 are blank)
                cells["A1"].PutValue("Header");
                cells["A2"].PutValue("Data1");
                cells["A4"].PutValue("Data2");
                cells["A6"].PutValue("Data3");

                // Set DeleteOptions to update references after deletion
                DeleteOptions options = new DeleteOptions
                {
                    UpdateReference = true
                };

                // Delete all blank rows on the first worksheet using the options
                cells.DeleteBlankRows(options);

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedBlankRows.xlsx");

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteBlankRowsWithUpdateReference.Run();
        }
    }
}
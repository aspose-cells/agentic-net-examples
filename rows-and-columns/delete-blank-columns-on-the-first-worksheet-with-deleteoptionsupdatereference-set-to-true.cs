using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DeleteBlankColumnsWithUpdateReference
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

                // Add sample data with a blank column (column B will be blank)
                cells["A1"].PutValue("Column A");
                cells["C1"].PutValue("Column C");
                cells["A2"].PutValue(1);
                cells["C2"].PutValue(3);

                // Set up DeleteOptions with UpdateReference = true
                DeleteOptions options = new DeleteOptions
                {
                    UpdateReference = true
                };

                // Delete all blank columns on the first worksheet
                cells.DeleteBlankColumns(options);

                // Save the workbook
                string outputPath = "DeletedBlankColumns.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            DeleteBlankColumnsWithUpdateReference.Run();
        }
    }
}
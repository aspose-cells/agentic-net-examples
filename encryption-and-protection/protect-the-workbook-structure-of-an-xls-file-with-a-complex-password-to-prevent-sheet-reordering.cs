using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class ProtectWorkbookStructure
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                using (Workbook workbook = new Workbook())
                {
                    // Add some data to the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Demo data");

                    // Define a complex password
                    string complexPassword = "P@55w0rd!#2023$%^&*()_+|~`";

                    // Protect only the workbook structure (prevents sheet reordering, addition, deletion)
                    workbook.Protect(ProtectionType.Structure, complexPassword);

                    // Define output file path
                    string outputPath = "ProtectedStructureWorkbook.xlsx";

                    // Save the protected workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                }
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as needed
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectWorkbookStructure.Run();
        }
    }
}
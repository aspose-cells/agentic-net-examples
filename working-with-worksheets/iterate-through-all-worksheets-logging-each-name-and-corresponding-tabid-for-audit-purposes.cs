using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAuditDemo
{
    public class WorksheetAudit
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // lifecycle: create

                // Add sample worksheets for demonstration
                workbook.Worksheets[0].Name = "FirstSheet";
                workbook.Worksheets.Add("SecondSheet");
                workbook.Worksheets.Add("ThirdSheet");

                // Iterate through all worksheets and log Name and TabId
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // TabId is the internal identifier of the sheet
                    Console.WriteLine($"Worksheet Name: {sheet.Name}, TabId: {sheet.TabId}");
                }

                // Save the workbook (optional, demonstrates lifecycle: save)
                string outputPath = "AuditDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error in WorksheetAudit.Run: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                WorksheetAudit.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
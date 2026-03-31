using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RetrieveWorksheetByNameDemo
    {
        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add a worksheet and give it a specific name
            Worksheet sheet = workbook.Worksheets.Add("DataSheet");
            sheet.Cells["A1"].PutValue("Original Sheet");

            // Retrieve the worksheet by its name using the indexer
            Worksheet retrievedSheet = workbook.Worksheets["DataSheet"];

            // Directly interact with the retrieved worksheet (e.g., write current date)
            retrievedSheet.Cells["B2"].PutValue(DateTime.Now);

            // Save the workbook to verify changes
            workbook.Save("RetrieveWorksheetByNameDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveWorksheetByNameDemo.Run();
        }
    }
}
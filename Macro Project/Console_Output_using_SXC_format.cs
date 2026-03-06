using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SaveAsSxcDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);

            // Display default save format (for reference)
            Console.WriteLine("Default SaveFormat: " + workbook.GetType().Name);

            // Save the workbook in StarOffice Calc (SXC) format
            string outputPath = "SampleOutput.sxc";
            workbook.Save(outputPath, SaveFormat.Sxc);

            // Confirm successful save
            Console.WriteLine($"Workbook saved successfully as SXC at: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SaveAsSxcDemo.Run();
        }
    }
}
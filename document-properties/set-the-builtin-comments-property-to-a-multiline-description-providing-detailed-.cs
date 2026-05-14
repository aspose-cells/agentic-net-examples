using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetBuiltInCommentsDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the built‑in document properties collection
            var properties = workbook.BuiltInDocumentProperties;

            // Set the Comments property to a multiline description
            properties.Comments = $"This workbook was generated automatically.{Environment.NewLine}" +
                                 $"It contains sample data for demonstration purposes.{Environment.NewLine}" +
                                 $"Please review the content and adjust as needed.{Environment.NewLine}" +
                                 $"Generated on: {DateTime.Now:f}";

            // Optional: display the set comment in the console
            Console.WriteLine("Workbook Comments:");
            Console.WriteLine(properties.Comments);

            // Save the workbook
            workbook.Save("WorkbookWithComments.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetBuiltInCommentsDemo.Run();
        }
    }
}
using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ValidateTitleBeforeExport
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Example: set the title (could be set elsewhere in real scenario)
            workbook.BuiltInDocumentProperties.Title = "Sample Document";

            // Validate that the Title built‑in property is not empty
            string title = workbook.BuiltInDocumentProperties.Title;
            if (string.IsNullOrWhiteSpace(title))
            {
                throw new InvalidOperationException("The workbook's Title property must not be empty before exporting.");
            }

            // Export the workbook to an external format (lifecycle save rule)
            workbook.Save("ValidatedWorkbook.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateTitleBeforeExport.Run();
        }
    }
}
// Title: Aspose.Cells C# – Set Multiline Comments Built‑In Document Property for a Workbook
// Description: This example creates a new Workbook, accesses its BuiltInDocumentPropertyCollection, assigns a multiline comment using a C# verbatim string literal, displays the value, and saves the file as WorkbookWithComments.xlsx.
// Keywords: Aspose.Cells | C# | .NET | BuiltInDocumentPropertyCollection | Comments property | multiline string | verbatim string literal | set workbook comments | Excel document properties | Workbook.Save example
// Common Searches: how to set comments built‑in document property Aspose.Cells C# | multiline workbook comments property Aspose.Cells .NET | Aspose.Cells set built‑in document properties example | save workbook after adding comments property Aspose.Cells | read Comments property from Excel file using Aspose.Cells
// Developer Intent: Add a detailed, multiline description to the workbook’s Comments built‑in document property and persist the workbook.
// Use Cases: Embed generation notes or audit details directly in the Excel file without modifying worksheets. | Provide reviewer instructions or feedback requests inside the workbook metadata. | Store a simple change‑log or version description that downstream processes can read.
// AI Prompts: Generate C# code with Aspose.Cells to set the Comments built‑in document property using a multiline verbatim string and save the workbook. | Show how to retrieve and display the Comments property after the workbook has been saved with Aspose.Cells. | Explain how line breaks are preserved in the Comments property when using a verbatim string literal in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    // This example creates a new Workbook, accesses its BuiltInDocumentPropertyCollection, assigns a multiline comment using a C# verbatim string literal, displays the value, and saves the file as WorkbookWithComments.xlsx.
    public class SetWorkbookComments
    {
        // Entry point required by the project
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the built‑in document properties collection
                BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

                // Set the Comments property with a multiline description
                properties.Comments = @"This workbook was generated programmatically.
It contains sample data for demonstration purposes.
Please review the content and provide feedback.";

                // Optionally display the set comment to verify
                Console.WriteLine("Workbook Comments:");
                Console.WriteLine(properties.Comments);

                // Save the workbook (lifecycle rule: save)
                string outputPath = "WorkbookWithComments.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

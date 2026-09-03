// Title: Add a multiline Comments built‑in document property to an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate a new Workbook, compose a multiline string that includes the current date, assign it to the built‑in Comments property, and save the file as an .xlsx. | Open an existing .xlsx file, update its Comments built‑in document property with a dynamic multi‑line note, and rewrite the workbook using Aspose.Cells in C#. | Write C# code that writes a multi‑line description into the Comments built‑in property, saves the workbook, and then reads back the property to confirm the content.
// Common Searches: Aspose.Cells C# set multiline Comments built‑in document property with current date | how to add notes to Excel built‑in properties using Aspose.Cells .NET | update workbook Comments property programmatically in C# Aspose.Cells | store multi‑line description in Excel metadata Aspose.Cells example
// Tags: set built‑in Comments property Aspose.Cells | multiline workbook description .NET | dynamic date in Excel Comments property | Aspose.Cells document metadata usage | C# write Excel built‑in properties

using System;
using Aspose.Cells;

// // Creates a workbook, builds a multiline string with a timestamp, assigns it to the built‑in Comments document property, and saves the workbook as WorkbookWithComments.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Multiline description for the Comments built‑in property
            string comments = @"This workbook contains the quarterly sales data.
Generated on: " + DateTime.Now.ToString("yyyy-MM-dd") + @"
Please review the charts and pivot tables for insights.";

            // Set the built‑in Comments property (feature rule)
            workbook.BuiltInDocumentProperties["Comments"].Value = comments;

            // Save the workbook (lifecycle rule)
            workbook.Save("WorkbookWithComments.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

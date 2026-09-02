// Title: Add a 'ProcessedDate' custom document property with the current DateTime to an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx file, adds a custom document property named 'ProcessedDate' set to DateTime.Now, and saves the workbook. | Use Aspose.Cells to insert a DateTime.Now custom property called ProcessedDate into a workbook and output the modified file. | Create a .NET routine that updates an Excel file's metadata by adding a ProcessedDate property with the current timestamp via Aspose.Cells.
// Common Searches: aspocells add custom document property processeddate c# | how to set current timestamp as a custom property in Excel using Aspose.Cells .NET | C# example for adding DateTime custom property to existing workbook with Aspose.Cells | save Excel file after adding custom metadata Aspose.Cells | update workbook custom document properties programmatically in .NET
// Tags: custom document property insertion Aspose.Cells | DateTime metadata update Excel .NET | programmatic workbook metadata modification | save workbook after custom property change | ProcessedDate Excel custom property

using System;
using Aspose.Cells;

// Loads an existing Excel workbook, adds a custom document property named 'ProcessedDate' with the current DateTime, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Add a custom document property named "ProcessedDate" with the current date and time
        workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

        // Save the workbook with the new property
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}

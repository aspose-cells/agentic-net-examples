// Title: Add a "ProcessedDate" custom document property with the current DateTime to an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add custom document property with current timestamp | how to set a ProcessedDate property in an Excel file using Aspose.Cells | C# code to add datetime custom property to workbook with Aspose.Cells | saving custom metadata to Excel workbook using Aspose.Cells .NET
// Tags: add custom document property Aspose.Cells | set datetime property workbook .NET | Aspose.Cells custom metadata Excel | save workbook with new property C# | process date custom property Excel Aspose

using System;
using Aspose.Cells;

// The example loads an existing Excel workbook, adds a custom document property named "ProcessedDate" containing the current DateTime, and saves the workbook to a new file.
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

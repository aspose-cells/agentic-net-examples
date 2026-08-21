// Title: C# – Add a custom ContentTypeProperty (ProjectId) to an Aspose.Cells workbook
// Description: Creates a new Workbook, uses the ContentTypeProperties.Add(name, value) overload to insert a string property called ProjectId (e.g., "12345"), and saves the file as ProjectWorkbook.xlsx.
// Keywords: Aspose.Cells ContentTypeProperties | C# add custom workbook property | Excel custom metadata Aspose | ProjectId workbook tag | set string property Aspose.Cells | Excel file metadata C#
// Common Searches: Aspose.Cells add custom ContentTypeProperty C# | How to store ProjectId in Excel with Aspose.Cells | C# add string metadata to workbook using Aspose | Save custom property in Excel file Aspose.Cells | Add and read ContentTypeProperties in C#
// Developer Intent: Insert a string‑typed ContentTypeProperty named ProjectId into a workbook’s metadata collection.
// Use Cases: Embed a unique project identifier for audit trails. | Tag exported reports with business‑specific codes. | Pass configuration values to downstream automation scripts via workbook metadata.
// AI Prompts: Write C# code that reads the ProjectId ContentTypeProperty from an existing Excel file using Aspose.Cells. | Show how to add several string ContentTypeProperties (e.g., ProjectId, Department) and persist them in a workbook. | Generate a snippet that sets ProjectId from a variable and includes error handling for missing properties.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates a new Workbook, uses the ContentTypeProperties.Add(name, value) overload to insert a string property called ProjectId (e.g., "12345"), and saves the file as ProjectWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a ContentTypeProperty named "ProjectId" with a string value
        // Using the overload that accepts name and value (type defaults to string)
        workbook.ContentTypeProperties.Add("ProjectId", "12345");

        // Save the workbook to a file
        workbook.Save("ProjectWorkbook.xlsx");
    }
}

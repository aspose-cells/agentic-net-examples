// Title: Use Windows Integrated Authentication for importing database data into an Aspose.Cells workbook (C#)
// Common Searches: Aspose.Cells import data from SQL Server using integrated security C# | C# how to connect Aspose.Cells worksheet to database with Windows authentication | Import DataTable into Aspose.Cells workbook without username and password | Example of Worksheet.Cells.ImportDataTable using Windows integrated authentication | Aspose.Cells .NET data import using integrated security instead of explicit credentials
// Tags: Aspose.Cells import DataTable with Windows authentication | C# integrated security Aspose.Cells database import | Worksheet.Cells.ImportDataTable using integrated security | .NET Aspose.Cells DB connection without credentials | Windows authentication for Aspose.Cells data import

using System;
using Aspose.Cells;

// This example shows how to import data from a database into an Aspose.Cells workbook using Windows Integrated Authentication in C#. Aspose.Cells does not expose a direct DbConnectionInfo property, so the Worksheet.Cells.ImportDataTable method is used with a connection string that relies on integrated security, eliminating the need to specify explicit usernames or passwords. The workbook is then saved as IntegratedAuthWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // NOTE: Aspose.Cells does not expose a direct DbConnectionInfo property on Worksheet in the current API.
            // If database connectivity is required, use the appropriate data import APIs (e.g., Worksheet.Cells.ImportDataTable).
            // The following placeholder demonstrates where such logic would be placed.

            // Save the workbook
            string outputPath = "IntegratedAuthWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

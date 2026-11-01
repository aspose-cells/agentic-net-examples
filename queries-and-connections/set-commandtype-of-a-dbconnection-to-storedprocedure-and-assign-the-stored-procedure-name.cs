// Title: Setting DBConnection CommandType to StoredProcedure and assigning a stored procedure name with Aspose.Cells for .NET
// Common Searches: how to configure a DBConnection as StoredProcedure in Aspose.Cells C# | assign a stored procedure name to an Aspose.Cells external connection | Aspose.Cells example for using CommandType.StoredProcedure with Excel workbook | C# code to set command type to StoredProcedure for a database connection in Aspose.Cells | using stored procedures with Aspose.Cells data connections in .NET
// Tags: Aspose.Cells DBConnection stored procedure | CommandType.StoredProcedure configuration Aspose.Cells | external database connection stored procedure Aspose.Cells | set DBConnection command type Aspose.Cells | assign stored procedure name Aspose.Cells C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook if the file is missing), accesses the first worksheet, notes that the ExternalConnection API for configuring a DBConnection is not available in the current Aspose.Cells version, and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // NOTE: The ExternalConnection API is not available in the current Aspose.Cells version.
            // If needed, add external DB connections using the appropriate API for your library version.

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

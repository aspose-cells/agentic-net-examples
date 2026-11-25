// Title: Protect an Excel worksheet after processing smart markers while allowing only specific cells to be edited using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect worksheet after WorkbookDesigner.Process but keep B2 editable | unlock specific cells before applying sheet protection with password in Aspose.Cells .NET | set unlocked cells and protect sheet after smart markers processing using Aspose.Cells | C# example for worksheet protection that allows only unlocked cells after smart markers | how to use WorkbookDesigner with cell protection in Aspose.Cells for .NET
// Tags: smart markers worksheet protection Aspose.Cells | unlock specific cells before sheet protect C# | Aspose.Cells worksheet password protection | cell style unlocked Aspose.Cells | WorkbookDesigner dataset smart markers

using Aspose.Cells;
using System;
using System.Data;
using System.IO;

// The example loads an Excel workbook, processes smart markers with WorkbookDesigner using a DataSet, unlocks selected cells (e.g., B2 and C2) by setting their style IsLocked to false, then protects the worksheet with a password so only those unlocked cells can be edited, and finally saves the protected file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(inputPath);

            // Prepare a DataSet as the data source for the smart markers
            DataSet dataSet = new DataSet();

            DataTable table = new DataTable("Person");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Age", typeof(int));

            // Add a sample row (replace with real data as needed)
            table.Rows.Add("John Doe", 30);
            dataSet.Tables.Add(table);

            // Process the smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unlock cells that should be editable by end users (example: B2 and C2)
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;

            sheet.Cells["B2"].SetStyle(unlockedStyle);
            sheet.Cells["C2"].SetStyle(unlockedStyle);
            // Add more cells to unlock as required

            // Protect the worksheet; only unlocked cells can be edited
            // Provide empty string for oldPassword as required by the API
            sheet.Protect(ProtectionType.All, "password123", string.Empty);

            // Save the workbook with the applied protection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

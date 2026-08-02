// Title: C# – Embed CSV as OLE Object with Auto‑Load in Excel using Aspose.Cells
// Description: Loads an existing workbook, inserts an OLE placeholder, embeds a CSV file as binary data, sets the ProgID "Excel.CSV" so Excel opens it, enables AutoLoad for automatic activation, and saves the updated file.
// Keywords: Aspose.Cells | C# | embed CSV | OLE object | ProgID | AutoLoad | Excel | OleObjects.Add | load workbook | embed file in Excel
// Common Searches: embed csv in excel using aspose.cells c# | add ole object to worksheet c# | set progid for ole object asp.net | auto load ole object aspose.cells | c# code to embed csv as ole object
// Developer Intent: Add a CSV file as an OLE object to an existing Excel workbook and configure it to open automatically with Excel.
// Use Cases: Create a single Excel package that contains source CSV data for easy reference by end users. | Distribute a template where the embedded CSV opens instantly when the workbook is opened, removing the need for separate data files. | Programmatically insert multiple CSV OLE objects across worksheets to consolidate raw data within one file.
// AI Prompts: Write C# code with Aspose.Cells to embed a CSV file as an OLE object, assign a custom icon, and set ProgID to Excel.CSV. | Explain how to read, modify, and replace the CSV data stored in an OLE object inside an existing workbook using Aspose.Cells. | Provide a step‑by‑step guide to add several CSV OLE objects to different worksheets and ensure each auto‑loads when the workbook is opened.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing workbook, inserts an OLE placeholder, embeds a CSV file as binary data, sets the ProgID "Excel.CSV" so Excel opens it, enables AutoLoad for automatic activation, and saves the updated file.
class AddCsvOleObject
{
    static void Main()
    {
        try
        {
            // Path to the existing workbook to load
            string workbookPath = "input.xlsx";

            // Path to the CSV file that will be embedded as an OLE object
            string csvFilePath = "data.csv";

            // Verify that the required files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");
            if (!File.Exists(csvFilePath))
                throw new FileNotFoundException($"CSV file not found: {csvFilePath}");

            // Load the workbook from file
            Workbook workbook = new Workbook(workbookPath);

            // Get the first worksheet (you can choose any worksheet as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Add an OLE object placeholder.
            // Passing null for imageData uses the default icon.
            int oleIndex = sheet.OleObjects.Add(5, 5, 200, 200, null);

            // Retrieve the newly added OLE object
            OleObject ole = sheet.OleObjects[oleIndex];

            // Set the embedded CSV data
            ole.ObjectData = File.ReadAllBytes(csvFilePath);

            // Set the ProgID so that the host application (Excel) knows how to open the CSV
            ole.ProgID = "Excel.CSV";

            // Configure the OLE object to load automatically when the workbook is opened
            ole.AutoLoad = true;

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

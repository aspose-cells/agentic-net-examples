using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Define custom Ribbon XML with a button that would invoke a macro named OnSelectSource
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"Data\">" +
            "        <group id=\"customGroup\" label=\"External Data\">" +
            "          <button id=\"btnSelectSource\" label=\"Select Source File\" size=\"large\" onAction=\"OnSelectSource\"/>" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        workbook.RibbonXml = ribbonXml;

        // Simulate the button click: ask user to input a file path
        string selectedFile = PromptFilePath();

        if (!string.IsNullOrEmpty(selectedFile))
        {
            // Access the first data connection (if any) and set its source file path via the connection string
            if (workbook.DataConnections.Count > 0)
            {
                var connection = workbook.DataConnections[0];
                // For external workbook connections the connection string format is: Data Source='full_path';
                connection.ConnectionString = $"Data Source='{selectedFile}';";
                Console.WriteLine($"Data connection connection string set to: {connection.ConnectionString}");
            }
            else
            {
                Console.WriteLine("No data connections found in the workbook.");
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Helper method that prompts the user to enter a file path
    static string PromptFilePath()
    {
        Console.Write("Enter the full path of the external data source file (or leave empty to cancel): ");
        string path = Console.ReadLine();
        return string.IsNullOrWhiteSpace(path) ? null : path.Trim();
    }
}
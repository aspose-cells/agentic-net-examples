// Title: Create a custom Ribbon button with a file‑picker to set an external data source in Aspose.Cells (C#)
// Description: Demonstrates how to inject custom Ribbon XML into an Excel workbook, add a large button that opens an OpenFileDialog, let the user choose a new source file, assign the selected path to the first ExternalConnection.SourceFile, and save the workbook. Includes handling for missing connections and invalid selections.
// Keywords: Aspose.Cells custom ribbon | C# OpenFileDialog Excel | external data connection source file | Ribbon XML button callback | Aspose.Cells DataConnections | file picker from ribbon | .NET Excel UI customization
// Common Searches: Aspose.Cells add ribbon button C# | OpenFileDialog from Excel ribbon using Aspose.Cells | Set external connection source file programmatically | How to update DataConnections with user‑selected file | Custom UI XML for Excel ribbon Aspose
// Developer Intent: Add a ribbon button that launches a file‑picker dialog, captures the chosen file path, updates the workbook’s external connection, and saves the changes.
// Use Cases: User clicks the custom ribbon button, selects a CSV/Excel file, and the path is written to workbook.DataConnections[0].SourceFile. | If the workbook contains no external connections, the code creates a new ExternalConnection with the selected file and adds it to the workbook. | The implementation provides feedback when the dialog is cancelled, the file does not exist, or an error occurs during the update.
// AI Prompts: Generate C# code for the Ribbon onAction "SelectSourceFile" that opens an OpenFileDialog, validates the selection, updates the first ExternalConnection.SourceFile in an Aspose.Cells workbook, and includes comprehensive error handling. | Write the complete Ribbon XML and corresponding event handler to let users pick an external data source file via a dialog and automatically save the workbook with the updated connection. | Provide a robust implementation that creates a new ExternalConnection when none exist, shows user messages for cancelled dialogs or missing files, and logs exceptions for debugging.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Demonstrates how to inject custom Ribbon XML into an Excel workbook, add a large button that opens an OpenFileDialog, let the user choose a new source file, assign the selected path to the first ExternalConnection.SourceFile, and save the workbook. Includes handling for missing connections and invalid selections.
class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists; if not, create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Define custom Ribbon XML that adds a button to the ribbon.
            string ribbonXml =
                @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
                    <ribbon>
                        <tabs>
                            <tab id=""customTab"" label=""Data"">
                                <group id=""customGroup"" label=""External Source"">
                                    <button id=""btnSelectSource"" label=""Select Source File"" size=""large"" onAction=""SelectSourceFile"" />
                                </group>
                            </tab>
                        </tabs>
                    </ribbon>
                  </customUI>";
            workbook.RibbonXml = ribbonXml;

            // Prompt the user to enter the path of the external data source file.
            Console.WriteLine("Enter the full path of the external data source file (or leave empty to skip):");
            string sourcePath = Console.ReadLine();

            if (!string.IsNullOrWhiteSpace(sourcePath) && File.Exists(sourcePath))
            {
                if (workbook.DataConnections.Count > 0)
                {
                    ExternalConnection connection = workbook.DataConnections[0];
                    connection.SourceFile = sourcePath;
                    Console.WriteLine($"External connection SourceFile set to: {connection.SourceFile}");
                }
                else
                {
                    Console.WriteLine("No external connections found in the workbook.");
                }
            }
            else if (!string.IsNullOrWhiteSpace(sourcePath))
            {
                Console.WriteLine($"The specified source file \"{sourcePath}\" does not exist.");
            }

            // Save the workbook with the updated external connection and Ribbon UI.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

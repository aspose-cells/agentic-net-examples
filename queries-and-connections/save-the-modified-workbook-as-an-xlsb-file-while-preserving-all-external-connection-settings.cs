// Title: Save workbook as XLSB while retaining external connections with Aspose.Cells for .NET
// Description: Loads an existing workbook, enables the SaveData flag on each DataConnection, configures XlsbSaveOptions (ExportAllColumnIndexes), and saves the file as XLSB, ensuring all external connection metadata remains intact.
// Keywords: Aspose.Cells | XLSB export | external data connections | SaveData property | DataConnections | XlsbSaveOptions | C# | .NET | convert xlsx to xlsb | preserve connection metadata | Power Query | workbook conversion
// Common Searches: Aspose.Cells save as XLSB with connections | C# preserve external data links when exporting to XLSB | Set SaveData for DataConnections before XLSB save | XlsbSaveOptions ExportAllColumnIndexes example | Convert .xlsx to .xlsb without losing Power Query links
// Developer Intent: Save an Excel workbook in XLSB format while keeping all external connection settings functional.
// Use Cases: Distribute a compact XLSB version of a reporting workbook that still refreshes Power Query sources. | Archive Excel files with linked data sources, preserving connection information for future updates. | Automate batch conversion of multiple .xlsx files that contain external connections to XLSB using Aspose.Cells.
// AI Prompts: Generate C# code with Aspose.Cells that loads an .xlsx, sets SaveData = true for each DataConnection, and saves as .xlsb with ExportAllColumnIndexes enabled. | Explain the impact of the SaveData property on external connections during XLSB export in Aspose.Cells. | Show how to configure XlsbSaveOptions to retain column indexes and external connection metadata when converting to XLSB.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsXlsbSaveExample
{
    // Loads an existing workbook, enables the SaveData flag on each DataConnection, configures XlsbSaveOptions (ExportAllColumnIndexes), and saves the file as XLSB, ensuring all external connection metadata remains intact.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your source file if needed)
            Workbook workbook = new Workbook("input.xlsx");

            // Preserve all external connection settings by ensuring the data fetched
            // over each connection is saved with the workbook.
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                // Keep the existing setting or explicitly enable it
                conn.SaveData = true;
            }

            // Create XLSB save options (default constructor as per the rule)
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                // Export all column indexes (default is true, set explicitly for clarity)
                ExportAllColumnIndexes = true
            };

            // Save the workbook as an XLSB file while preserving external connections
            workbook.Save("output.xlsb", saveOptions);
        }
    }
}

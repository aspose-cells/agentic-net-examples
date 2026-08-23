// Title: How to compare OData DataModelConnection version metadata before and after modification using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, adds a DataModelConnection with a versioned OData URL, saves the file, reloads it, updates the version in the connection string, and checks the version consistency with Aspose.Cells. | Write a C# method that extracts the Version parameter from a DataModelConnection.ConnectionString and determines whether the original and updated versions match after the workbook is saved.
// Common Searches: aspnet compare OData connection version after editing workbook with Aspose.Cells | extract version from DataModelConnection connection string C# | validate external OData connection metadata consistency in Excel file using Aspose.Cells | how to update OData service version in Aspose.Cells workbook programmatically | check if OData connection version changed after saving workbook Aspose.Cells .NET
// Tags: DataModelConnection version extraction C# | modify external OData connection string Aspose.Cells | verify OData version consistency after workbook save | parse connection string parameters Aspose.Cells .NET | track OData service version changes in Excel file

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsODataMetadataComparison
{
    // Shows how to create a DataModelConnection that simulates an OData source, save the workbook, reload it, change the version in the connection string, and programmatically extract and compare the original and updated version values to detect any mismatches.
    class Program
    {
        static void Main()
        {
            // ---------- Create ----------
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Create a DataModelConnection (used here to simulate an OData connection)
            // The class does not have a public constructor, so we use reflection as shown in Aspose examples
            DataModelConnection odataConnection = (DataModelConnection)System.Runtime.Serialization.FormatterServices
                .GetUninitializedObject(typeof(DataModelConnection));

            // Set an initial connection string that includes a version identifier
            // (In real OData connections, the version might be part of the URL or a query parameter)
            odataConnection.ConnectionString = "Provider=MSOLAP;Data Source=https://service.example.com/odata;Version=1.0";

            // Assign a unique name for the connection
            odataConnection.Name = "SampleODataConnection";

            // Add the connection to the workbook's data connections collection
            ((IList<ExternalConnection>)workbook.DataConnections).Add(odataConnection);

            // Save the workbook (initial state)
            workbook.Save("ODataDemo_before.xlsx");

            // ---------- Load ----------
            // Load the workbook we just saved
            Workbook loadedWorkbook = new Workbook("ODataDemo_before.xlsx");

            // Retrieve the first external connection and cast it to DataModelConnection
            DataModelConnection loadedConnection = loadedWorkbook.DataConnections[0] as DataModelConnection;
            if (loadedConnection == null)
            {
                Console.WriteLine("No OData (DataModel) connection found.");
                return;
            }

            // Capture metadata before modification
            string originalConnectionString = loadedConnection.ConnectionString;
            Console.WriteLine("Original ConnectionString: " + originalConnectionString);

            // ---------- Modify ----------
            // Update the connection string to a new version (simulating a version change)
            loadedConnection.ConnectionString = "Provider=MSOLAP;Data Source=https://service.example.com/odata;Version=2.0";

            // Save the workbook after modification
            loadedWorkbook.Save("ODataDemo_after.xlsx");

            // ---------- Load Again ----------
            // Load the modified workbook to verify changes
            Workbook modifiedWorkbook = new Workbook("ODataDemo_after.xlsx");
            DataModelConnection modifiedConnection = modifiedWorkbook.DataConnections[0] as DataModelConnection;
            if (modifiedConnection == null)
            {
                Console.WriteLine("No OData (DataModel) connection found after modification.");
                return;
            }

            // Capture metadata after modification
            string updatedConnectionString = modifiedConnection.ConnectionString;
            Console.WriteLine("Updated ConnectionString: " + updatedConnectionString);

            // ---------- Compare ----------
            // Simple version extraction assuming the format "...;Version=X.Y"
            string GetVersion(string connStr)
            {
                const string versionKey = "Version=";
                int start = connStr.IndexOf(versionKey, StringComparison.OrdinalIgnoreCase);
                if (start < 0) return string.Empty;
                start += versionKey.Length;
                int end = connStr.IndexOf(';', start);
                return end > start ? connStr.Substring(start, end - start) : connStr.Substring(start);
            }

            string originalVersion = GetVersion(originalConnectionString);
            string updatedVersion = GetVersion(updatedConnectionString);

            Console.WriteLine($"Original Version: {originalVersion}");
            Console.WriteLine($"Updated Version: {updatedVersion}");

            if (originalVersion == updatedVersion)
            {
                Console.WriteLine("Version is consistent (no change detected).");
            }
            else
            {
                Console.WriteLine("Version mismatch detected. Ensure version consistency before proceeding.");
            }
        }
    }
}

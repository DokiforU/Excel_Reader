# Excel_Reader

![C#](https://img.shields.io/badge/language-C%23-blue.svg) ![.NET](https://img.shields.io/badge/framework-.NET-purple.svg) ![License](https://img.shields.io/badge/license-MIT-green.svg)

This is a simple console application developed using C# and the [ClosedXML](https://github.com/ClosedXML/ClosedXML) library. Its core function is to read content from Excel (.xlsx) files and efficiently convert it into `System.Data.DataTable` objects, which is highly convenient for subsequent data processing or database import operations.

---

## ✨ Features

*   **Read Specific Worksheet**: Accurately reads data from a worksheet by its name.
*   **Default Reading**: Automatically reads the first worksheet if no name is specified.
*   **Iterate All Worksheets**: Reads data from all worksheets in an Excel file with a single command.
*   **Custom Header Row**: Flexibly specify which row contains the column headers.
*   **Duplicate Column Name Handling**: Automatically detects and renames duplicate column names to ensure the robustness of the `DataTable`.
*   **Cross-Platform**: Built on .NET, it can run on Windows, macOS, and Linux.

---

## 🚀 How to Use

1.  **Clone the Repository**
    ```bash
    git clone git@github.com:DokiforU/Excel_Reader.git
    cd Excel_Reader
    ```

2.  **Restore Dependencies**
    This project uses `packages.config` to manage NuGet dependencies.
    *   **Using an IDE**: Open the solution file (`test_7.12_c#.sln`) with Visual Studio or JetBrains Rider. The IDE will automatically prompt you to restore the NuGet packages.
    *   **Using the Command Line**: If you have `nuget.exe` configured, you can run `nuget restore`. For projects based on the .NET SDK, `dotnet restore` is typically used.

3.  **Prepare Your Excel File**
    Place the `.xlsx` file you want to read into the project directory, or use an absolute/relative path.

4.  **Modify the Code to Specify Your File**
    Open `Program.cs` and modify the `fileName` variable in the `Main` method to point to your Excel file.

    ```csharp
    // Program.cs
    class Program
    {
        static void Main(string[] args)
        {
            // Please replace this with the path to your Excel file
            string fileName = "your_file_name.xlsx"; 
            
            // ... rest of the code
        }
    }
    ```

5.  **Run the Project**
    *   Click the "Run" button in Visual Studio or Rider.
    *   Alternatively, execute `dotnet run` in the terminal (if the project is a .NET Core/5+ type).

---

## 📝 Code Examples

Using the `ExcelDataReader` class is very straightforward.

```csharp
using System;
using System.Data;
using ExcelReader; // Import your namespace

class Example
{
    static void Main(string[] args)
    {
        string filePath = "path/to/your/file.xlsx";

        try
        {
            // Use 'using' to ensure resources are properly disposed
            using (var reader = new ExcelDataReader(filePath))
            {
                // Example 1: Read the first worksheet
                Console.WriteLine("--- Reading the first worksheet ---");
                DataTable firstSheet = reader.ReadSheet();
                PrintDataTable(firstSheet);

                // Example 2: Read a worksheet by name
                Console.WriteLine("\n--- Reading worksheet named 'Sheet1' ---");
                DataTable specificSheet = reader.ReadSheet("Sheet1");
                PrintDataTable(specificSheet);

                // Example 3: Read all worksheets
                Console.WriteLine("\n--- Reading all worksheets ---");
                var allSheets = reader.ReadAllSheets();
                foreach (var sheet in allSheets)
                {
                    Console.WriteLine($"\nWorksheet: {sheet.Key}");
                    PrintDataTable(sheet.Value);
                }
                
                // Example 4: Specify the header is on row 3
                using (var readerWithCustomHeader = new ExcelDataReader(filePath, headerRow: 3))
                {
                    Console.WriteLine("\n--- Reading sheet with header on row 3 ---");
                    DataTable customHeaderSheet = readerWithCustomHeader.ReadSheet();
                    PrintDataTable(customHeaderSheet);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // (Your PrintDataTable method...)
}
```

---

## 📁 Project Structure

*   `ExcelDataReader.cs`: The core class that encapsulates all Excel reading logic.
*   `Program.cs`: The program's entry point and an example of how to use `ExcelDataReader`.
*   `test_7.12_c#.sln` / `test_7.12_c#.csproj`: Visual Studio solution and project files.
*   `packages.config`: The list of NuGet dependencies for the project.

---

## 🔧 Prerequisites

*   .NET Framework / .NET Core
*   [ClosedXML](https://github.com/ClosedXML/ClosedXML) (v0.105.0 or higher)

---

## 💡 Future Plans

*   [ ] Add functionality to write data to a new Excel file.
*   [ ] Add support for the older Excel (`.xls`) format (might require other libraries).
*   [ ] Add unit tests to ensure code stability and reliability.

---

## 👨‍💻 Author

*   **yuli_xia** - [DokiforU](https://github.com/DokiforU)

---

## 📜 License

This project is licensed under the MIT License. See the `LICENSE` file for details.

# FriendlyExcel

[![CI](https://github.com/Vladoss46/FriendlyExcel/actions/workflows/ci.yml/badge.svg)](https://github.com/Vladoss46/FriendlyExcel/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE.txt)

.NET 8 library for reading and writing Excel files (`.xls` / `.xlsx`) through `System.Data.DataTable`.

Built on [NPOI](https://github.com/nissl-lab/npoi).

## Install

```bash
dotnet add package FriendlyExcel
```

Requires **.NET 8**.

## Read Excel → DataTable

```csharp
using FriendlyExcel;
using FriendlyExcel.Models;
using System.Data;

// Path + index / name
DataTable table = ExcelReader.Get("data.xlsx");
DataTable byIndex = ExcelReader.Get("data.xlsx", sheetIndex: 1);
DataTable byName = ExcelReader.Get("data.xlsx", "Orders");

// Options object
DataTable typed = ExcelReader.Get("data.xlsx", new ExcelReadOptions
{
    SheetName = "Orders",
    UseFirstRowAsColumnNames = true,
    ColumnTypes = [typeof(string), typeof(int), typeof(double)],
});

// Stream (web uploads, blobs) — Format or FileName required
await using Stream upload = file.OpenReadStream();
DataTable fromUpload = ExcelReader.Get(upload, new ExcelReadOptions
{
    FileName = file.FileName, // e.g. "report.xlsx"
    SheetName = "Sheet1",
});

// Whole workbook
XLBook book = ExcelReader.GetBook("data.xlsx", true, true);
XLBook fromStream = ExcelReader.GetBook(upload, new ExcelReadBookOptions
{
    Format = ExcelFormat.Xlsx,
    UseFirstRowAsColumnNames = true,
});
```

Column types are inferred when not provided: `string`, `int`, `double`, `bool`, `DateTime`, `DateOnly`, `TimeOnly`.  
`DataTable.TableName` is set from the worksheet name. Empty sheets become empty tables when reading a whole book.

Unicode text (including Russian / Cyrillic) in sheet names, headers, and cells is preserved.

**Numbers and Russian Excel:** in the Russian UI the decimal separator is a **comma** (`12,5`).  
Inside the file, real numeric cells are still binary doubles (locale-agnostic).  
Text cells like `"12,5"` / `"12.5"` are parsed flexibly (comma or point) into `int`/`double` when the column is numeric — without the InvariantCulture pitfall where `"12,5"` becomes `125`.

## Write DataTable → Excel

```csharp
using FriendlyExcel;
using FriendlyExcel.Models;
using System.Data;

DataTable people = new("People");
people.Columns.Add("Name", typeof(string));
people.Columns.Add("Age", typeof(int));
people.Rows.Add("Alice", 30);

DataTable orders = new("Orders");
orders.Columns.Add("Id", typeof(int));
orders.Columns.Add("Total", typeof(double));
orders.Rows.Add(1, 99.5);

// Path — format from extension
ExcelWriter.Save("out.xlsx", people);

// Stream — Format required
using MemoryStream ms = new();
ExcelWriter.Save(ms, people, ExcelFormat.Xlsx);

XLBook book = new([people, orders]);
ms.SetLength(0);
ExcelWriter.SaveBook(ms, book, new ExcelWriteOptions
{
    Format = ExcelFormat.Xlsx,
    WriteColumnNames = true,
});
```

Sheet names come from `DataTable.TableName` when set.

## Read / write POCOs

Map columns to properties by name (case-insensitive), or with `[ExcelColumn]`:

```csharp
public sealed class Employee
{
    [ExcelColumn("Имя")]
    public string Name { get; set; } = "";

    [ExcelColumn("Зарплата")]
    public double Salary { get; set; }

    public bool Active { get; set; } // column "Active"
}

List<Employee> people = ExcelReader.Get<Employee>("staff.xlsx");
ExcelWriter.Save("staff-out.xlsx", people);

using MemoryStream ms = new();
ExcelWriter.Save(ms, people, ExcelFormat.Xlsx);
ms.Position = 0;
List<Employee> again = ExcelReader.Get<Employee>(ms, new ExcelReadOptions { Format = ExcelFormat.Xlsx });
```

## Public API

| Type | Role |
|------|------|
| `ExcelReader` | Read path/stream → `DataTable`, `List<T>`, or `XLBook` |
| `ExcelWriter` | Write `DataTable`, `IEnumerable<T>`, or `XLBook` → path/stream |
| `ExcelColumnAttribute` | Map a property to an Excel column name |
| `ExcelReadOptions` / `ExcelReadBookOptions` | Sheet index/name, headers, types, format |
| `ExcelWriteOptions` | Headers, format |
| `ExcelFormat` | `Xls` / `Xlsx` |
| `XLBook` | List of sheets as `DataTable`s |
| `EmptySheetException` | Thrown for empty worksheets (single-sheet read) |

## Build & test

```bash
dotnet build
dotnet test
```

## License

FriendlyExcel is licensed under the [MIT License](LICENSE.txt).

Third-party dependencies (NPOI and its transitive packages) are listed in [NOTICE](NOTICE). NPOI is Apache-2.0; when you redistribute binaries that include NPOI, follow Apache-2.0 attribution requirements for that dependency.

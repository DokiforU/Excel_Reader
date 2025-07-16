// ExcelDataReader.cs

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Data;

namespace ExcelReader
{
    public class ExcelDataReader : IDisposable
    {
        private readonly XLWorkbook _workbook;
        private int _headerRow; // 表头行
        private int _dataStartRow; // 数据起始行

        public ExcelDataReader(string filePath, int headerRow = 1)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("文件未找到。", filePath);
            }

            _workbook = new XLWorkbook(filePath);
            _headerRow = headerRow;
            _dataStartRow = headerRow + 1;
        }

        // 根据工作表名称读取数据
        public DataTable ReadSheet(string sheetName = null)
        {
            IXLWorksheet worksheet;

            if (string.IsNullOrEmpty(sheetName))
            {
                // 默认读取第一个工作表
                worksheet = _workbook.Worksheet(1);
                Console.WriteLine($"未指定工作表名称，读取第一个工作表: '{worksheet.Name}'");
            }
            else
            {
                // 根据名称查找工作表
                if (!_workbook.Worksheets.TryGetWorksheet(sheetName, out worksheet))
                {
                    throw new ArgumentException($"名为 '{sheetName}' 的工作表不存在。");
                }
            }

            return ReadDataFromWorksheet(worksheet);
        }

        // 读取所有工作表
        public Dictionary<string, DataTable> ReadAllSheets()
        {
            var allSheetsData = new Dictionary<string, DataTable>();
            foreach (var worksheet in _workbook.Worksheets)
            {
                Console.WriteLine($"--- 处理工作表: {worksheet.Name} ---");
                var table = ReadDataFromWorksheet(worksheet);
                if (table.Rows.Count > 0)
                {
                    allSheetsData.Add(worksheet.Name, table);
                }
            }
            return allSheetsData;
        }

        // 从指定的工作表对象中提取数据到DataTable
        private DataTable ReadDataFromWorksheet(IXLWorksheet worksheet)
        {
            var table = new DataTable(worksheet.Name);

            var headerRow = worksheet.Row(_headerRow);
            if (headerRow.IsEmpty())
            {
                Console.WriteLine($"工作表 '{worksheet.Name}' 的第 {_headerRow} 行（预期为表头）为空，返回空表。");
                return table;
            }

            var headers = headerRow.CellsUsed().Select(cell => cell.Value.ToString().Trim()).ToList(); // 生成一个只包含表头字符串的集合
            foreach (var header in headers)
            {
                if (!table.Columns.Contains(header))
                {
                    table.Columns.Add(header, typeof(object));
                }
                else
                {
                    table.Columns.Add($"{header}_{Guid.NewGuid().ToString().Substring(0, 4)}", typeof(object));
                    Console.WriteLine($"工作表 '{worksheet.Name}' 发现重复列名 '{header}'，已自动重命名。");
                }
            }

            var dataRows = worksheet.RowsUsed().Where(r => r.RowNumber() >= _dataStartRow); // 过滤出行号 >= 数据起始行号的行

            // 遍历所有数据行
            foreach (var dataRow in dataRows)
            {
                DataRow row = table.NewRow();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    row[i] = dataRow.Cell(i + 1).Value; // 把Excel单元格的值,赋给DataRow中对应位置的单元格
                }
                table.Rows.Add(row);
            }

            return table;
        }

        public void Dispose()
        {
            _workbook?.Dispose();
        }
    }
}
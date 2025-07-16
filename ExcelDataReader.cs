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
        private int _headerRow; // ��ͷ��
        private int _dataStartRow; // ������ʼ��

        public ExcelDataReader(string filePath, int headerRow = 1)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("�ļ�δ�ҵ���", filePath);
            }

            _workbook = new XLWorkbook(filePath);
            _headerRow = headerRow;
            _dataStartRow = headerRow + 1;
        }

        // ���ݹ��������ƶ�ȡ����
        public DataTable ReadSheet(string sheetName = null)
        {
            IXLWorksheet worksheet;

            if (string.IsNullOrEmpty(sheetName))
            {
                // Ĭ�϶�ȡ��һ��������
                worksheet = _workbook.Worksheet(1);
                Console.WriteLine($"δָ�����������ƣ���ȡ��һ��������: '{worksheet.Name}'");
            }
            else
            {
                // �������Ʋ��ҹ�����
                if (!_workbook.Worksheets.TryGetWorksheet(sheetName, out worksheet))
                {
                    throw new ArgumentException($"��Ϊ '{sheetName}' �Ĺ��������ڡ�");
                }
            }

            return ReadDataFromWorksheet(worksheet);
        }

        // ��ȡ���й�����
        public Dictionary<string, DataTable> ReadAllSheets()
        {
            var allSheetsData = new Dictionary<string, DataTable>();
            foreach (var worksheet in _workbook.Worksheets)
            {
                Console.WriteLine($"--- ��������: {worksheet.Name} ---");
                var table = ReadDataFromWorksheet(worksheet);
                if (table.Rows.Count > 0)
                {
                    allSheetsData.Add(worksheet.Name, table);
                }
            }
            return allSheetsData;
        }

        // ��ָ���Ĺ������������ȡ���ݵ�DataTable
        private DataTable ReadDataFromWorksheet(IXLWorksheet worksheet)
        {
            var table = new DataTable(worksheet.Name);

            var headerRow = worksheet.Row(_headerRow);
            if (headerRow.IsEmpty())
            {
                Console.WriteLine($"������ '{worksheet.Name}' �ĵ� {_headerRow} �У�Ԥ��Ϊ��ͷ��Ϊ�գ����ؿձ�");
                return table;
            }

            var headers = headerRow.CellsUsed().Select(cell => cell.Value.ToString().Trim()).ToList(); // ����һ��ֻ������ͷ�ַ����ļ���
            foreach (var header in headers)
            {
                if (!table.Columns.Contains(header))
                {
                    table.Columns.Add(header, typeof(object));
                }
                else
                {
                    table.Columns.Add($"{header}_{Guid.NewGuid().ToString().Substring(0, 4)}", typeof(object));
                    Console.WriteLine($"������ '{worksheet.Name}' �����ظ����� '{header}'�����Զ���������");
                }
            }

            var dataRows = worksheet.RowsUsed().Where(r => r.RowNumber() >= _dataStartRow); // ���˳��к� >= ������ʼ�кŵ���

            // ��������������
            foreach (var dataRow in dataRows)
            {
                DataRow row = table.NewRow();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    row[i] = dataRow.Cell(i + 1).Value; // ��Excel��Ԫ���ֵ,����DataRow�ж�Ӧλ�õĵ�Ԫ��
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
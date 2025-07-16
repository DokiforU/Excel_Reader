// Program.cs

using System;
using System.Data;
using ExcelReader;

class Program
{
    static void Main(string[] args)
    {
        string fileName = ".xlsx";

        try
        {
            using (var reader = new ExcelDataReader(fileName))
            {
                // 读取第一个工作表
                DataTable table = reader.ReadSheet(); // 不传名字，默认读第一个
                PrintDataTable(table);
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n发生错误: {ex.Message}");
            Console.ResetColor();
        }

        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }

    // 打印DataTable
    public static void PrintDataTable(DataTable table)
        {
            if (table == null || table.Rows.Count == 0)
            {
                Console.WriteLine("表格为空。");
                return;
            }

            // 打印表头
            foreach (DataColumn col in table.Columns)
            {
                Console.Write($"{col.ColumnName,-20}"); // 左对齐，总宽度20
            }
            Console.WriteLine();
            Console.WriteLine(new string('-', table.Columns.Count * 20));

            // 打印数据行
            foreach (DataRow row in table.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write($"{item?.ToString(),-20}");
                }
                Console.WriteLine();
            }
        }
}
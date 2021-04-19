using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Data;
using ExcelDataReader;
using System.Windows.Forms;

namespace ExcelSheetListMaker
{
    class Program
    {
        [STAThreadAttribute]
        static void Main(string[] args)
        {
            Task task = ExecuteAsync(args);
            task.Wait();
        }

        static async Task ExecuteAsync(string[] args)
        {
            // コマンドライン引数からファイルパス、フォルダパスを取得する
            var files = args.Where(x => File.Exists(x)).ToList();
            var directories = args.Where(x => Directory.Exists(x)).ToList();
            // フォルダパスから全てのファイルパスを取得する
            var filesInSubdirecories = directories.SelectMany(x => Directory.EnumerateFiles(x, "*", SearchOption.AllDirectories)).ToList();
            files.AddRange(filesInSubdirecories);

            // 該当する拡張子のファイルパスを抽出する
            string[] excelExtentions = { ".xlsx", ".xlsm", ".xlsb", ".xls", ".xls" };
            var excelFiles = files.Where(x => !Path.GetFileName(x).StartsWith("~") && excelExtentions.Contains(Path.GetExtension(x), StringComparer.OrdinalIgnoreCase));

            if (!excelFiles.Any())
            {
                return;
            }


            ExcelData[] excelDatas = await Task.WhenAll(excelFiles.OrderBy(x => x).Select(ReadExcelDataAsinc));

            StringBuilder sb = new StringBuilder();
            sb.Append("パス\tフォルダ\tファイル\tシート\r\n");

            foreach (var excelData in excelDatas)
            {
                if (excelData.Sheets == null)
                {
                    continue;
                }

                foreach (DataTable tbl in excelData.Sheets.Tables)
                {
                    sb.Append(excelData.path);
                    sb.Append("\t");
                    sb.Append(Path.GetDirectoryName(excelData.path));
                    sb.Append("\t");
                    sb.Append(Path.GetFileName(excelData.path));
                    sb.Append("\t");
                    sb.Append(tbl.TableName);
                    sb.Append("\r\n");
                }
            }

            Clipboard.SetText(sb.ToString());
        }


        static async Task<ExcelData> ReadExcelDataAsinc(string path)
        {
            return await Task.Run<ExcelData>(() => ReadExcelData(path));
        }


        static ExcelData ReadExcelData(string path)
        {
            DataSet ds = null;
            try
            {
                using(var stream = File.Open(path, FileMode.Open, FileAccess.Read)){
                    using(var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        ds = reader.AsDataSet();
                    }
                }

            }
            catch(Exception ex)
            {
                Console.Error.WriteLine(path);
                Console.Error.WriteLine(ex.ToString());
                Console.Read();
            }

            return new ExcelData { path = path, Sheets = ds };
        }
    }

    class ExcelData
    {
        public string path;
        public DataSet Sheets;
    }
}

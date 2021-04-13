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
            // コマンドライン引数からファイルパス、フォルダパスを取得する
            var files = args.Where(x => File.Exists(x)).ToList();
            var directories = args.Where(x => Directory.Exists(x)).ToList();
            // フォルダパスから全てのファイルパスを取得する
            var filesInSubdirecories = directories.SelectMany(x => Directory.EnumerateFiles(x, "*", SearchOption.AllDirectories)).ToList();
            files.AddRange(filesInSubdirecories);

            // 該当する拡張子のファイルパスを抽出する
            string[] excelExtentions = { ".xlsx", ".xlsm", ".xlsb", ".xls", ".xls" };
            var excelFiles = files.Where(x => !Path.GetFileName(x).StartsWith("~") && excelExtentions.Contains(Path.GetExtension(x), StringComparer.OrdinalIgnoreCase));

            if(!excelFiles.Any())
            {
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("パス\tフォルダ\tファイル\tシート\r\n");

            foreach (var file in excelFiles.OrderBy(x => x))
            {
                var ds = ReadExcelData(file);
                if(ds == null)
                {
                    continue;
                }

                foreach(DataTable tbl in ds.Tables)
                {
                    sb.Append(file);
                    sb.Append("\t");
                    sb.Append(Path.GetDirectoryName(file));
                    sb.Append("\t");
                    sb.Append(Path.GetFileName(file));
                    sb.Append("\t");
                    sb.Append(tbl.TableName);
                    sb.Append("\r\n");
                }
            }

            Clipboard.SetText(sb.ToString());          
        }

        static DataSet ReadExcelData(string path)
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
            }

            return ds;
        }
    }
}

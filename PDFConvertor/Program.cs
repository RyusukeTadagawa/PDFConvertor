using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NLog;
using NetOffice.OfficeApi;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.ExcelApi.Enums;

namespace PDFConvertor
{
    class Program
    {
        /// <summary>
        /// Nlogクラスインスタンス
        /// </summary>
        /// <returns></returns>
        private static Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 当プログラムの戻り値
        /// </summary>
        public enum ReturnCode
        {

            /// <summary>
            /// 正常終了
            /// </summary>
            Normal,
            /// <summary>
            /// 引数エラー
            /// </summary>
            ParameterError,
            /// <summary>
            /// 例外により処理終了
            /// </summary>
            ExceptionOccur
        }

        /// <summary>
        /// メインメソッド
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static int Main(string[] args)
        {
            try
            {
                var ExePath = Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
                var nlog_config_path = Path.Combine(ExePath, "NLog.config");
                LogManager.Configuration = new NLog.Config.XmlLoggingConfiguration(nlog_config_path);
                LogManager.Configuration.Variables["ExePath"] = ExePath;

                if(!CheckArgs(args))
                {
                    return (int)ReturnCode.ParameterError;
                }

                var excelPath = args[0];
                var outputDir = ExePath;
                if(args.Length > 1)
                {
                    outputDir = args[1];
                }

                ConvertExcel2Pdf(excelPath, outputDir);
                return (int)ReturnCode.Normal;

            }
            catch (Exception ex)
            {
                _logger.Fatal(ex);
                return (int)ReturnCode.ExceptionOccur;
            }
            finally
            {
                LogManager.Shutdown();
            }


        }

        /// <summary>
        /// 指定のExcelファイルの全シートを指定のフォルダにPDFに変換する
        /// </summary>
        /// <param name="excelPath">Excelファイルパス</param>
        /// <param name="outputDir">出力フォルダパス</param>
        private static void ConvertExcel2Pdf(string excelPath, string outputDir)
        {
            // start excel and turn off msg boxes
            Application excelApplication = new Application();
            excelApplication.DisplayAlerts = false;


            // add a new workbook
            Workbook workBook = excelApplication.Workbooks.Open(excelPath);

            int sheetCount = workBook.Sheets.Count;
            for (int i =1; i <= sheetCount;i++)
            {
                Worksheet worksheet = workBook.Worksheets[i] as Worksheet;
                //Console.WriteLine(worksheet.Name);
                var pdfPath = Path.Combine(outputDir, worksheet.Name + ".pdf");
                //Console.WriteLine(pdfPath);
                worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfPath, XlFixedFormatQuality.xlQualityStandard);
            }


            workBook.Close();

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();
        }


        /// <summary>
        /// プログラム引数のチェックメソッド
        /// </summary>
        /// <param name="args">プログラム引数</param>
        /// <returns>チェックの可否</returns>
        private static Boolean CheckArgs(string[] args)
        {
            Boolean wblnReturn = false;
            if (args.Length < 1)
            {
                _logger.Error("引数の指定が不正です。");
                return wblnReturn;
            }

            var excelFilepath = args[0];

            if (!File.Exists(excelFilepath))
            {
                _logger.Error("指定のファイルを確認できません。");
                return wblnReturn;
            }

            if(args.Length > 1)
            {
                var outputDir = args[1];
                if(!Directory.Exists(outputDir))
                {
                    _logger.Error("指定のディレクトリを確認できません。");
                    return wblnReturn;
                }
            }

            wblnReturn = true;
            return wblnReturn;

        }


    }
}

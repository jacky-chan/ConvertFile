using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using System.Web;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;

namespace ConvertSwf
{
    public class Convert2Swf
    {
        // Fields
        private static XlFixedFormatType excelType;
        private static PpSaveAsFileType ppType;
        private static WdExportFormat wd;

        //构造函数
        public Convert2Swf()
        {
        }
        /// <summary>
        /// pdf2swf转换
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <param name="pdf2swfStr"></param>
        private static void ConvertCmd(string sourcePath, string targetPath, string pdf2swfStr)
        {
            try
            {
                using (Process process = new Process())
                {
                    //pdf2swf.exe的调用路径
                    string path = HttpContext.Current.Server.MapPath(pdf2swfStr);
                    if (!File.Exists(path))
                    {
                        throw new ApplicationException("Can not find: " + path);
                    }
                    //转换命令
                    string arguments = "  -t " + sourcePath + " -s flashversion=9 -o " + targetPath;
                    //新建进程
                    ProcessStartInfo info = new ProcessStartInfo(path, arguments);
                    process.StartInfo = info;
                    info.CreateNoWindow = true;
                    //隐藏命令窗口
                    info.WindowStyle = ProcessWindowStyle.Hidden;
                    //开始进程
                    process.Start();
                    //等待完成退出
                    process.WaitForExit();
                    //如果没有转后成功，加上图片转为位图（较慢，图片会失真）
                    if (!File.Exists(targetPath))
                    {
                        arguments = arguments + " -s poly2bitmap";
                        ProcessStartInfo info2 = new ProcessStartInfo(path, arguments);
                        process.StartInfo = info2;
                        info2.CreateNoWindow = true;
                        //隐藏命令窗口
                        info2.WindowStyle = ProcessWindowStyle.Hidden;
                        //开始进程
                        process.Start();
                        //等待完成退出
                        process.WaitForExit();
                    }
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }

        }
        /// <summary>
        /// 转换文件入口
        /// </summary>
        /// <param name="filePath">需要转换的文件本地路径</param>
        /// <param name="webName">网站根目录名称</param>
        /// <param name="pdf2swfStr">pdf2swf.exe路径</param>
        /// <param name="jpg2swfStr">jpg2swf.exe路径（无用）</param>
        /// <returns>转换后返回的.swf路径</returns>
        public static string ConvertFile(string filePath, string webName, string pdf2swfStr, string jpg2swfStr)
        {
            string str5;
            string str6;
            string str7;
            string str = "";
            string path = filePath.Substring(0, filePath.LastIndexOf(".")) + ".swf";
            //如果存在已经转换的文件不再转换，直接返回
            if (File.Exists(path))
            {
                return LocalDirToWebDir(path, webName);
            }
            string str3 = filePath.Substring(filePath.LastIndexOf(".") + 1);
            string str4 = filePath.Substring(filePath.LastIndexOf(@"\") + 1);
            if ("swf".Equals(str3))
            {
                return LocalDirToWebDir(filePath, webName);
            }
            //转换pdf
            if ("pdf".Equals(str3))
            {
                str5 = filePath;
                str6 = filePath.Substring(0, filePath.LastIndexOf(".")) + ".swf";
                ConvertCmd(str5, str6, pdf2swfStr);
                return LocalDirToWebDir(str6, webName);
            }
            //转换word
            if ("doc".Equals(str3) || "docx".Equals(str3))
            {
                str5 = filePath;
                str6 = filePath.Substring(0, filePath.LastIndexOf(".")) + ".pdf";
                if (ConvertOffice2PDF(str5, str6, wd))
                {
                    str7 = filePath.Substring(0, filePath.LastIndexOf(".")) + ".swf";
                    ConvertCmd(str6, str7, pdf2swfStr);
                    str = LocalDirToWebDir(str7, webName);
                    File.Delete(str6);
                }
                return str;
            }
            //转换excel
            if ("xlsx".Equals(str3) || "xls".Equals(str3))
            {
                str5 = filePath;
                str6 = filePath.Substring(0, filePath.LastIndexOf(".")) + ".pdf";
                if (ConvertOffice2PDF(str5, str6, excelType))
                {
                    str7 = filePath.Substring(0, filePath.LastIndexOf(".")) + ".swf";
                    ConvertCmd(str6, str7, pdf2swfStr);
                    str = LocalDirToWebDir(str7, webName);
                    File.Delete(str6);
                }
                return str;
            }
            //转换ppt
            if ("ppt".Equals(str3) || "pptx".Equals(str3))
            {
                str5 = filePath;
                str6 = filePath.Substring(0, filePath.LastIndexOf(".")) + ".pdf";
                if (ConvertOffice2PDF(str5, str6, ppType))
                {
                    str7 = filePath.Substring(0, filePath.LastIndexOf(".")) + ".swf";
                    ConvertCmd(str6, str7, pdf2swfStr);
                    str = LocalDirToWebDir(str7, webName);
                    File.Delete(str6);
                }
                return str;
            }
            return "0";

        }
        /// <summary>
        /// jpg转swf
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <param name="jpg2swfStr"></param>
        private static void ConvertJpgCmd(string sourcePath, string targetPath, string jpg2swfStr)
        {
            try
            {
                using (Process process = new Process())
                {
                    string path = HttpContext.Current.Server.MapPath(jpg2swfStr);
                    if (!File.Exists(path))
                    {
                        throw new ApplicationException("Can not find: " + path);
                    }
                    string arguments = "  -o " + targetPath + " " + sourcePath + " -s flashversion=9";
                    ProcessStartInfo info = new ProcessStartInfo(path, arguments);
                    process.StartInfo = info;
                    info.CreateNoWindow = true;
                    info.WindowStyle = ProcessWindowStyle.Hidden;
                    process.Start();
                    process.WaitForExit();
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }

        }
        /// <summary>
        /// Excel转换为pdf
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <param name="targetType"></param>
        /// <returns></returns>
        private static bool ConvertOffice2PDF(string sourcePath, string targetPath, XlFixedFormatType targetType)
        {
            bool result;
            object missing = Type.Missing;
            Microsoft.Office.Interop.Excel.ApplicationClass application = null;
            Workbook workBook = null;
            try
            {
                application = new Microsoft.Office.Interop.Excel.ApplicationClass();
                object target = targetPath;
                object type = targetType;
                workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);

                workBook.ExportAsFixedFormat(targetType, target, XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (workBook != null)
                {
                    workBook.Close(true, missing, missing);
                    workBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;

        }
        /// <summary>
        /// PPt转换为pdf
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <param name="targetFileType"></param>
        /// <returns></returns>
        private static bool ConvertOffice2PDF(string sourcePath, string targetPath, PpSaveAsFileType targetFileType)
        {
            bool result;
            object missing = Type.Missing;
            Microsoft.Office.Interop.PowerPoint.ApplicationClass application = null;
            Presentation persentation = null;
            try
            {
                application = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(targetPath, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue);

                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;

        }
        /// <summary>
        /// 转换word为pdf
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <param name="exportFormat"></param>
        /// <returns></returns>
        private static bool ConvertOffice2PDF(string sourcePath, string targetPath, WdExportFormat exportFormat)
        {
            bool result;
            object paramMissing = Type.Missing;
            Microsoft.Office.Interop.Word.ApplicationClass wordApplication = new Microsoft.Office.Interop.Word.ApplicationClass();
            Microsoft.Office.Interop.Word.Document wordDocument = null;
            try
            {
                object paramSourceDocPath = sourcePath;
                string paramExportFilePath = targetPath;

                Microsoft.Office.Interop.Word.WdExportFormat paramExportFormat = exportFormat;
                bool paramOpenAfterExport = false;
                Microsoft.Office.Interop.Word.WdExportOptimizeFor paramExportOptimizeFor =
                        Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                Microsoft.Office.Interop.Word.WdExportRange paramExportRange = Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                Microsoft.Office.Interop.Word.WdExportItem paramExportItem = Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                Microsoft.Office.Interop.Word.WdExportCreateBookmarks paramCreateBookmarks =
                        Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;

                wordDocument = wordApplication.Documents.Open(
                        ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing);

                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                            paramExportFormat, paramOpenAfterExport,
                            paramExportOptimizeFor, paramExportRange, paramStartPage,
                            paramEndPage, paramExportItem, paramIncludeDocProps,
                            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                            paramBitmapMissingFonts, paramUseISO19005_1,
                            ref paramMissing);
                result = true;
            }
            finally
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        /// <summary>
        /// 本地路径转换成网络路径
        /// </summary>
        /// <param name="LDir"></param>
        /// <param name="WebName"></param>
        /// <returns></returns>
        private static string LocalDirToWebDir(string LDir, string WebName)
        {
            string[] strArray = LDir.Split(new char[] { Convert.ToChar(@"\") });
            int index = 0;
            do
            {
                index++;
            }
            while (strArray[index] != WebName);
            string str = "";
            for (int i = index + 1; i < (strArray.Length - 1); i++)
            {
                str = str + strArray[i] + "/";
            }
            str = str + strArray[strArray.Length - 1];
            return ("/" + WebName + "/" + str);

        }
    }


}

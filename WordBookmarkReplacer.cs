using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using wordbookmark2pdf.handlers;
using OpenXmlBookmarkStart = DocumentFormat.OpenXml.Wordprocessing.BookmarkStart;

namespace wordbookmark2pdf
{
    /// <summary>
    /// 书签模板替换器
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class WordBookmarkReplacer<T>
    {
        private readonly IWordBookmarkHandler<T> _handler;
        private readonly IDictionary<string, MethodInfo> _handlerMethods
            = new ConcurrentDictionary<string, MethodInfo>();

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="handler">书签模板处理器</param>
        public WordBookmarkReplacer(IWordBookmarkHandler<T> handler)
        {
            _handler = handler;

            foreach (var methodInfo in _handler.GetType().GetMethods())
            {
                var parameterInfos = methodInfo.GetParameters();
                if (parameterInfos.Length == 3)
                {
                    if (parameterInfos[0].ParameterType == typeof(Text)
                        || parameterInfos[0].ParameterType == typeof(ImagePart)
                        || parameterInfos[0].ParameterType == typeof(DocumentFormat.OpenXml.Wordprocessing.Table)
                        && parameterInfos[1].ParameterType == typeof(ReplacerContext<T>)
                        && parameterInfos[2].ParameterType == typeof(Action<string>))
                    {
                        _handlerMethods.Add(methodInfo.Name.ToUpper(), methodInfo);
                    }
                }
            }
        }

        /// <summary>
        /// 开始处理模板替换
        /// </summary>
        /// <param name="ctx">上下文</param>
        /// <param name="progress">进度汇报</param>
        /// <param name="logger">日志记录</param>
        /// <param name="missMatch">未命中书签事件</param>
        public void Handle(ReplacerContext<T> ctx, ProgressChangedEventHandler progress,
            Action<string> logger, Action<string> missMatch)
        {
            var filepath = ctx.SrcFilePath;
            var targetFolder = ctx.DestFolder;
            if (!File.Exists(filepath))
            {
                logger?.Invoke($"{filepath}不存在");
                return;
            }

            var name = Path.GetFileName(filepath);
            if (string.IsNullOrEmpty(name))
            {
                logger?.Invoke($"文件名称{filepath}为空");
                return;
            }

            var targetPath = ctx.DestFilePath;
            try
            {
                File.Copy(filepath, targetPath, true);
            }
            catch (Exception e)
            {
                logger?.Invoke($"拷贝模板文件失败:{e.Message}");
                return;
            }
            progress?.Invoke(this, new ProgressChangedEventArgs(10, "预处理完毕"));
            using (var stream = File.Open(targetPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (var wordprocessingDocument = WordprocessingDocument.Open(stream, true))
                {
                    var mainDocumentPart = wordprocessingDocument.MainDocumentPart;

                    var bookmarkStarts = mainDocumentPart.Document.Body.Descendants<OpenXmlBookmarkStart>().AsQueryable();
                    var bookMarkConfig = ctx.EnsureBookMarkConf(logger);// 加载配置文件
                    var count = bookmarkStarts.Count();
                    int c = 0;
                    DocumentFormat.OpenXml.Wordprocessing.Table table;
                    foreach (var bookmarkStart in bookmarkStarts)
                    {
                        ++c;
                        progress?.Invoke(this, new ProgressChangedEventArgs(10 + 70 * c / count, $"开始替换书签<{bookmarkStart.Name}>"));
                        ctx.BookmarkName = bookmarkStart.Name;

                        if (bookmarkStart.Name.Value.ToUpper().Equals("_GOBACK"))
                        {
                            logger?.Invoke($"书签[{bookmarkStart.Name.Value}]:处理成表格书签");
                            continue;
                        }

                        if (!_handlerMethods.ContainsKey(bookmarkStart.Name.Value.ToUpper()))
                        {
                            // 处理未知书签
                            UpdateBookMarksByCfg(bookmarkStart, bookMarkConfig, ctx, logger, missMatch);
                            continue;
                        }

                        if (CheckParentTable(bookmarkStart, out table) && null != table)
                        {
                            // 处理表格
                            logger?.Invoke($"书签[{bookmarkStart.Name.Value}]:处理成表格书签");
                            _handlerMethods[bookmarkStart.Name.Value.ToUpper()].Invoke(_handler,
                                new object[] { table, ctx, logger });

                            continue;
                        }

                        var bookmarkText = bookmarkStart.NextSibling<Run>();
                        if (null != bookmarkText)
                        {
                            var firstChild = bookmarkText.GetFirstChild<Text>();
                            if (firstChild != null)
                            {
                                // 处理文本
                                logger?.Invoke($"书签[{bookmarkStart.Name.Value}]:处理成文本书签");
                                _handlerMethods[bookmarkStart.Name.Value.ToUpper()].Invoke(_handler,
                                    new object[] { firstChild, ctx, logger });

                                continue;
                            }

                            ImagePart imagePart = null;
                            if (CheckPicture(bookmarkStart, ref imagePart))
                            {
                                // 处理图片
                                logger?.Invoke($"书签[{bookmarkStart.Name.Value}]:处理成图片书签");
                                _handlerMethods[bookmarkStart.Name.Value.ToUpper()].Invoke(_handler,
                                    new object[] { imagePart, ctx, logger });
                                continue;
                            }
                        }

                        logger?.Invoke($"书签[{bookmarkStart.Name.Value}]:不是预期类型的书签");
                    }

                    wordprocessingDocument.Close();
                }
            }

            if (!Directory.Exists(targetFolder))
            {
                Directory.CreateDirectory(targetFolder);
            }

            try
            {
                progress?.Invoke(this, new ProgressChangedEventArgs(80, "开始pdf转换"));
                Convert(targetPath, Path.Combine(targetFolder, ctx.SrcFileNameWithoutExt + ".pdf"));
                progress?.Invoke(this, new ProgressChangedEventArgs(90, "pdf转换成功"));
            }
            catch (Exception e)
            {
                logger?.Invoke($"转换pdf失败:{e.Message}");
            }
        }

        private bool CheckParentTable(OpenXmlBookmarkStart bookmarkstart, out DocumentFormat.OpenXml.Wordprocessing.Table table)
        {
            if (bookmarkstart.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Table)
            {
                table = bookmarkstart.NextSibling() as DocumentFormat.OpenXml.Wordprocessing.Table;
                return true;
            }

            var parent = bookmarkstart.Parent;
            while (parent != null)
            {
                if (parent is DocumentFormat.OpenXml.Wordprocessing.Table)
                {
                    table = parent as DocumentFormat.OpenXml.Wordprocessing.Table;
                    return true;
                }

                parent = parent.Parent;
            }

            table = null;
            return false;
        }

        private bool CheckPicture(OpenXmlBookmarkStart bookmarkstart, ref ImagePart imagePart)
        {
            var document = bookmarkstart.Parent;

            while (document != null)
            {
                if (document is DocumentFormat.OpenXml.Wordprocessing.Document)
                {
                    break;
                }

                document = document.Parent;
            }

            if (document == null) return false;

            var doc = document as DocumentFormat.OpenXml.Wordprocessing.Document;
            var mainpart = doc.MainDocumentPart;
            var run = bookmarkstart.NextSibling<Run>();
            var pic = run?.GetFirstChild<Drawing>();
            if (pic != null)
            {
                var blipElement = pic.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                if (blipElement != null)
                {
                    var imageId = blipElement.Embed.Value;
                    imagePart = (ImagePart)mainpart.GetPartById(imageId);
                }

                return true;
            }

            return false;
        }

        private void Convert(string fileFullName, string tgr)
        {
            if (File.Exists(fileFullName))
            {
                var appWord = new Application();
                var wordDocument = appWord.Documents.Open(fileFullName);
                try
                {
                    wordDocument.ExportAsFixedFormat(tgr, WdExportFormat.wdExportFormatPDF);
                }
                finally
                {
                    wordDocument.Close();
                    appWord.Quit();
                }
            }
        }

        /// <summary>
        /// xml配置文件中添加的书签替换
        /// </summary>
        /// <param name="bookmarkstart"></param>
        /// <param name="bookMarkConfs"></param>
        /// <param name="ctx"></param>
        /// <param name="logger"></param>
        /// <param name="missMatch"></param>
        private bool UpdateBookMarksByCfg(OpenXmlBookmarkStart bookmarkstart, BookMarkConfs bookMarkConfs, ReplacerContext<T> ctx, Action<string> logger, Action<string> missMatch)
        {
            if (null == bookMarkConfs || bookMarkConfs.Confs.Count < 1)
            {
                logger?.Invoke("配置文件:宏替换器内容为空");
                return false;
            }

            if (!bookMarkConfs.Confs.ContainsKey(ctx.DestFileNameWithoutExt))
            {
                logger?.Invoke($"配置文件:没有找到{ctx.DestFileNameWithoutExt}宏替换器");
                return false;
            }

            var bookMarks = bookMarkConfs.Confs[ctx.DestFileNameWithoutExt].Marks;

            if (bookMarks.ContainsKey(bookmarkstart.Name.Value.ToUpper()))
            {
                var bookMark = bookMarks[bookmarkstart.Name.Value.ToUpper()];
                switch (bookMark.Type)
                {
                    case "TEXT":
                        {
                            ctx.BookmarkText = ctx.PreprocessText(bookMark.Text);
                            ctx.BookmarkSql = ctx.PreprocessSql(bookMark.Sql);

                            var bookmarkText = bookmarkstart.NextSibling<Run>();
                            var text = bookmarkText?.GetFirstChild<Text>();
                            if (text != null)
                            {
                                _handler.CustomText(text, ctx, logger);
                                return true;
                            }
                            else
                            {
                                logger?.Invoke($"书签[{bookmarkstart.Name.Value}]:不是文本书签");
                            }
                        }
                        break;
                    case "PICTURE":
                        {
                            ctx.BookmarkPicUri = ctx.PreprocessPicUri(bookMark.PicUri);
                            ImagePart imagepart = null;
                            if (CheckPicture(bookmarkstart, ref imagepart))
                            {
                                _handler.CustomPicture(imagepart, ctx, logger);
                                return true;
                            }
                            else
                            {
                                logger?.Invoke($"书签[{bookmarkstart.Name.Value}]:不是图片书签");
                            }
                        }
                        break;
                    case "TABLE":
                        {
                            ctx.BookmarkSql = ctx.PreprocessSql(bookMark.Sql);
                            DocumentFormat.OpenXml.Wordprocessing.Table table;
                            if (CheckParentTable(bookmarkstart, out table))
                            {
                                _handler.CustomTable(table, ctx, logger);
                                return true;
                            }
                            else
                            {
                                logger?.Invoke($"书签[{bookmarkstart.Name.Value}]:不是表格书签");
                            }
                        }
                        break;

                    default:
                        {
                            logger?.Invoke($"书签[{bookmarkstart.Name.Value}]:类型定义{bookMark.Type}不正确(TEXT/PICTURE/TABLE之一)");
                        }
                        return false;
                }

                logger?.Invoke($"书签[{bookmarkstart.Name.Value}]:替换完成");
            }
            else
            {
                logger?.Invoke($"书签[{bookmarkstart.Name.Value}]:没有找到宏替换器");
            }
            return false;
        }
    }
}

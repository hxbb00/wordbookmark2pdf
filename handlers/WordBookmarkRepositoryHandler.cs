using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace wordbookmark2pdf.handlers
{
    /// <summary>
    /// 模板处理器抽象基类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class WordBookmarkRepositoryHandler<T> : IWordBookmarkHandler<T>
    {
        /// <summary>
        /// 模板名称
        /// </summary>
        public abstract string TemplateName { get; }

        /// <summary>
        /// 自定义图片处理器
        /// </summary>
        /// <param name="t">图片</param>
        /// <param name="ctx">上下文</param>
        /// <param name="logger">日志记录器</param>
        /// <returns></returns>
        public bool CustomPicture(ImagePart t, ReplacerContext<T> ctx, Action<string> logger = null)
        {
            if (File.Exists(ctx.BookmarkPicUri))
            {
                using (var fs = new FileStream(ctx.BookmarkPicUri, FileMode.Open,
                    FileAccess.Read, FileShare.ReadWrite))
                {
                    t.FeedData(fs);
                }

                logger?.Invoke($"图片{ctx.BookmarkPicUri}替换成功");
                return true;
            }
            else
            {
                logger?.Invoke($"图片{ctx.BookmarkPicUri}不存在");
            }

            return false;
        }

        /// <summary>
        /// 自定义表格处理器
        /// </summary>
        /// <param name="t">表格对象</param>
        /// <param name="ctx">上下文</param>
        /// <param name="logger">日志记录器</param>
        /// <returns></returns>
        public bool CustomTable(Table t, ReplacerContext<T> ctx, Action<string> logger = null)
        {
            IDataReader reader = null;
            var dt = new DataTable();
            try
            {
                var command = ctx.CreateCommand();
                command.CommandText = ctx.BookmarkSql;
                reader = command.ExecuteReader();

                dt.Locale = CultureInfo.CurrentCulture;
                dt.Load(reader);
            }
            catch (Exception e)
            {
                var format = $"替换[{ctx.DestFilePath}]表格书签[{ctx.BookmarkName}]发生SQL({ctx.BookmarkSql})错误:\r\n{e.Message}";
                logger?.Invoke(format);
                Console.WriteLine(format);
                Console.WriteLine(e);
            }
            finally
            {
                reader?.Dispose();
            }

            if (dt.Rows.Count < 1) return false;

            var rowfld = t.GetFirstChild<TableRow>();
            t.RemoveAllChildren<TableRow>();

            var grid = t.GetFirstChild<TableGrid>();
            t.InsertAfter(rowfld, grid);

            var count = Math.Min(rowfld.ChildElements.Count(ce => ce is TableCell), dt.Columns.Count);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                TableRow row = new TableRow();

                for (int j = 0; j < count; j++)
                {
                    var cell = new TableCell();
                    cell.AppendChild(new Paragraph(new Run(new Text(dt.Rows[i][j].ToString()))));
                    row.AppendChild(cell);
                }

                if (t.LastChild is TableRow)
                {
                    t.InsertAfter(row, t.LastChild);
                }
                else
                {
                    t.InsertBefore(row, t.LastChild);
                }
            }
            return true;
        }

        /// <summary>
        /// 自定义文本处理器
        /// </summary>
        /// <param name="t">文本</param>
        /// <param name="ctx">上下文</param>
        /// <param name="logger">日志记录器</param>
        /// <returns></returns>
        public bool CustomText(Text t, ReplacerContext<T> ctx, Action<string> logger = null)
        {
            // 文本不为空,优先替换文本
            if (!string.IsNullOrEmpty(ctx.BookmarkText))
            {
                t.Text = ctx.BookmarkText;
                return false;
            }

            if (string.IsNullOrEmpty(ctx.BookmarkSql))
            {
                return false;
            }

            IDataReader reader = null;
            var dt = new DataTable();
            try
            {
                var command = ctx.CreateCommand();
                command.CommandText = ctx.BookmarkSql;
                reader = command.ExecuteReader();

                dt.Locale = CultureInfo.CurrentCulture;
                dt.Load(reader);
            }
            catch (Exception e)
            {
                var format = $"替换[{ctx.DestFilePath}]文本书签[{ctx.BookmarkName}]发生SQL({ctx.BookmarkSql})错误:\r\n{e.Message}";
                logger?.Invoke(format);
                Console.WriteLine(format);
                Console.WriteLine(e);
            }
            finally
            {
                reader?.Dispose();
            }

            if (dt.Columns.Count < 1 || dt.Rows.Count < 1) return false;

            t.Text = dt.Rows[0][0].ToString();
            return true;
        }
    }
}
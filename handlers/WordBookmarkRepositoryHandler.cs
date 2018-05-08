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
    /// ģ�崦�����������
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class WordBookmarkRepositoryHandler<T> : IWordBookmarkHandler<T>
    {
        /// <summary>
        /// ģ������
        /// </summary>
        public abstract string TemplateName { get; }

        /// <summary>
        /// �Զ���ͼƬ������
        /// </summary>
        /// <param name="t">ͼƬ</param>
        /// <param name="ctx">������</param>
        /// <param name="logger">��־��¼��</param>
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

                logger?.Invoke($"ͼƬ{ctx.BookmarkPicUri}�滻�ɹ�");
                return true;
            }
            else
            {
                logger?.Invoke($"ͼƬ{ctx.BookmarkPicUri}������");
            }

            return false;
        }

        /// <summary>
        /// �Զ���������
        /// </summary>
        /// <param name="t">������</param>
        /// <param name="ctx">������</param>
        /// <param name="logger">��־��¼��</param>
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
                var format = $"�滻[{ctx.DestFilePath}]�����ǩ[{ctx.BookmarkName}]����SQL({ctx.BookmarkSql})����:\r\n{e.Message}";
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
        /// �Զ����ı�������
        /// </summary>
        /// <param name="t">�ı�</param>
        /// <param name="ctx">������</param>
        /// <param name="logger">��־��¼��</param>
        /// <returns></returns>
        public bool CustomText(Text t, ReplacerContext<T> ctx, Action<string> logger = null)
        {
            // �ı���Ϊ��,�����滻�ı�
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
                var format = $"�滻[{ctx.DestFilePath}]�ı���ǩ[{ctx.BookmarkName}]����SQL({ctx.BookmarkSql})����:\r\n{e.Message}";
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
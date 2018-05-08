using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace wordbookmark2pdf
{
    /// <summary>
    /// �����Ļ���
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class ReplacerContext<T>
    {
        /// <summary>
        /// ���캯��
        /// </summary>
        /// <param name="filepath">ģ��·��</param>
        /// <param name="targetDir">����ļ���</param>
        protected ReplacerContext(string filepath, string targetDir)
        {
            DestFolder = targetDir;
            SrcFilePath = filepath;
        }

        /// <summary>
        /// �������Զ������
        /// </summary>
        public T Model { get; set; }
        /// <summary>
        /// ԭģ��·��
        /// </summary>
        public string SrcFilePath { get; private set; }
        /// <summary>
        /// Ŀ���ļ���
        /// </summary>
        public string DestFolder { get; private set; }
        /// <summary>
        /// ԭģ�岻����׺�ļ���
        /// </summary>
        public string SrcFileNameWithoutExt { get { return Path.GetFileNameWithoutExtension(SrcFilePath); } }
        /// <summary>
        /// ����ļ�������׺�ļ���
        /// </summary>
        public string DestFileNameWithoutExt { get { return Path.GetFileNameWithoutExtension(DestFilePath); } }
        /// <summary>
        /// ����ļ�·��
        /// </summary>
        public string DestFilePath { get { return Path.Combine(DestFolder, $"{Path.GetFileName(SrcFilePath)}"); } }
        /// <summary>
        /// ��ǩ����
        /// </summary>
        public string BookmarkName { get; protected internal set; }
        /// <summary>
        /// ��ǩͼƬ����
        /// </summary>
        public string BookmarkPicUri { get; protected internal set; }
        /// <summary>
        /// ��ǩsql
        /// </summary>
        public string BookmarkSql { get; protected internal set; }
        /// <summary>
        /// ��ǩ�ı�
        /// </summary>
        public string BookmarkText { get; protected internal set; }

        /// <summary>
        /// ��ǩ�����ļ�����
        /// </summary>
        /// <param name="logger">��־��¼��</param>
        /// <returns></returns>
        public virtual BookMarkConfs EnsureBookMarkConf(Action<string> logger)
        {
            string bookMarkConfig = BookMarkConfigPath;
            if (!File.Exists(bookMarkConfig))
            {
                logger?.Invoke($"�����ļ�:{bookMarkConfig}������");
                return null;
            }

            var xmlDoc = new XmlDocument();

            try
            {
                xmlDoc.Load(bookMarkConfig);
                logger?.Invoke($"�����ļ�:{bookMarkConfig}���سɹ�");
            }
            catch (Exception e)
            {
                logger?.Invoke(e.Message);
                Console.WriteLine(e);
                return null;
            }

            var bookMarkConfs = new BookMarkConfs();
            var doc = xmlDoc.DocumentElement;
            if (doc == null
                || doc.ChildNodes.Count == 0)
            {
                logger?.Invoke($"�����ļ�:{bookMarkConfig}�������κνڵ���Ϣ");
                return null;
            }

            foreach (XmlNode nodeMatch in doc.ChildNodes)
            {
                if (nodeMatch.Attributes == null) continue;
                if (nodeMatch.ChildNodes.Count > 0)
                {
                    var templName = nodeMatch.Attributes["name"].Value;

                    var bookMarks = new BookMarks(templName);
                    bookMarkConfs.Confs.Add(templName, bookMarks);

                    foreach (XmlNode node in nodeMatch.ChildNodes)
                    {
                        if (node.Attributes == null) continue;
                        var markname = node.Attributes["name"].Value.ToUpper();
                        var type = node.Attributes["type"].Value.ToUpper();

                        var bookMark = new BookMark(type);
                        bookMarks.Marks.Add(markname, bookMark);

                        bookMark.Text = node.Attributes["text"].Value;
                        bookMark.Sql = node.Attributes["sql"].Value;
                        bookMark.PicUri = node.Attributes["url"].Value;
                    }
                }
            }

            return bookMarkConfs;
        }

        /// <summary>
        /// ��ǩ�����ļ�·��
        /// </summary>
        protected virtual string BookMarkConfigPath
        {
            get { return "BookMarkConf.xml"; }
        }

        /// <summary>
        /// Ԥ������ǩ�ı�(���糣�ú�Ĵ���@SYS.TIME:'yyyyMMdd')
        /// </summary>
        /// <param name="text">��ǩ�ı�</param>
        /// <returns>��ǩ�ı�</returns>
        public virtual string PreprocessText(string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                var pattern1 = @"(?<=@SYS.TIME:').+?(?=')";
                foreach (Match match in Regex.Matches(text, pattern1))
                {
                    var fmt = match.Value;
                    var time = $"{DateTime.Now.ToString(fmt)}";
                    text = text.Replace($"@SYS.TIME:'{fmt}'", time);// ϵͳ��ǰʱ��
                }
            }

            return text;
        }

        /// <summary>
        /// Ԥ������ǩsql
        /// </summary>
        /// <param name="sql">��ǩsql</param>
        /// <returns>��ǩsql</returns>
        public virtual string PreprocessSql(string sql)
        {
            return sql;
        }

        /// <summary>
        /// Ԥ������ǩͼƬURL
        /// </summary>
        /// <param name="url">��ǩͼƬURL</param>
        /// <returns>��ǩͼƬURL</returns>
        public virtual string PreprocessPicUri(string url)
        {
            return url;
        }

        /// <summary>
        /// ��ǩsqlִ��ʹ�õ����ݿ�ִ������(��sqlִ�п��Է��ؿ�,����ʵ��)
        /// </summary>
        /// <returns></returns>
        public abstract IDbCommand CreateCommand();
    }
}
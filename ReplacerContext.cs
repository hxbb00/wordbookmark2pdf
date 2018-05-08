using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace wordbookmark2pdf
{
    /// <summary>
    /// 上下文基类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class ReplacerContext<T>
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filepath">模板路径</param>
        /// <param name="targetDir">存放文件夹</param>
        protected ReplacerContext(string filepath, string targetDir)
        {
            DestFolder = targetDir;
            SrcFilePath = filepath;
        }

        /// <summary>
        /// 上下文自定义对象
        /// </summary>
        public T Model { get; set; }
        /// <summary>
        /// 原模板路径
        /// </summary>
        public string SrcFilePath { get; private set; }
        /// <summary>
        /// 目标文件夹
        /// </summary>
        public string DestFolder { get; private set; }
        /// <summary>
        /// 原模板不带后缀文件名
        /// </summary>
        public string SrcFileNameWithoutExt { get { return Path.GetFileNameWithoutExtension(SrcFilePath); } }
        /// <summary>
        /// 结果文件不带后缀文件名
        /// </summary>
        public string DestFileNameWithoutExt { get { return Path.GetFileNameWithoutExtension(DestFilePath); } }
        /// <summary>
        /// 结果文件路径
        /// </summary>
        public string DestFilePath { get { return Path.Combine(DestFolder, $"{Path.GetFileName(SrcFilePath)}"); } }
        /// <summary>
        /// 书签名称
        /// </summary>
        public string BookmarkName { get; protected internal set; }
        /// <summary>
        /// 书签图片名称
        /// </summary>
        public string BookmarkPicUri { get; protected internal set; }
        /// <summary>
        /// 书签sql
        /// </summary>
        public string BookmarkSql { get; protected internal set; }
        /// <summary>
        /// 书签文本
        /// </summary>
        public string BookmarkText { get; protected internal set; }

        /// <summary>
        /// 书签配置文件解析
        /// </summary>
        /// <param name="logger">日志记录器</param>
        /// <returns></returns>
        public virtual BookMarkConfs EnsureBookMarkConf(Action<string> logger)
        {
            string bookMarkConfig = BookMarkConfigPath;
            if (!File.Exists(bookMarkConfig))
            {
                logger?.Invoke($"配置文件:{bookMarkConfig}不存在");
                return null;
            }

            var xmlDoc = new XmlDocument();

            try
            {
                xmlDoc.Load(bookMarkConfig);
                logger?.Invoke($"配置文件:{bookMarkConfig}加载成功");
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
                logger?.Invoke($"配置文件:{bookMarkConfig}不存在任何节点信息");
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
        /// 书签配置文件路径
        /// </summary>
        protected virtual string BookMarkConfigPath
        {
            get { return "BookMarkConf.xml"; }
        }

        /// <summary>
        /// 预处理书签文本(比如常用宏的处理@SYS.TIME:'yyyyMMdd')
        /// </summary>
        /// <param name="text">书签文本</param>
        /// <returns>书签文本</returns>
        public virtual string PreprocessText(string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                var pattern1 = @"(?<=@SYS.TIME:').+?(?=')";
                foreach (Match match in Regex.Matches(text, pattern1))
                {
                    var fmt = match.Value;
                    var time = $"{DateTime.Now.ToString(fmt)}";
                    text = text.Replace($"@SYS.TIME:'{fmt}'", time);// 系统当前时间
                }
            }

            return text;
        }

        /// <summary>
        /// 预处理书签sql
        /// </summary>
        /// <param name="sql">书签sql</param>
        /// <returns>书签sql</returns>
        public virtual string PreprocessSql(string sql)
        {
            return sql;
        }

        /// <summary>
        /// 预处理书签图片URL
        /// </summary>
        /// <param name="url">书签图片URL</param>
        /// <returns>书签图片URL</returns>
        public virtual string PreprocessPicUri(string url)
        {
            return url;
        }

        /// <summary>
        /// 书签sql执行使用的数据库执行命令(无sql执行可以返回空,建议实现)
        /// </summary>
        /// <returns></returns>
        public abstract IDbCommand CreateCommand();
    }
}
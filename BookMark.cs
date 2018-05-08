namespace wordbookmark2pdf
{
    /// <summary>
    /// 书签对象
    /// </summary>
    public class BookMark
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="type">书签类型</param>
        public BookMark(string type)
        {
            Type = type;
        }

        /// <summary>
        /// 书签类型
        /// </summary>
        public string Type { get; private set; }
        /// <summary>
        /// 书签文本
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// 书签sql
        /// </summary>
        public string Sql { get; set; }
        /// <summary>
        /// 书签图片
        /// </summary>
        public string PicUri { get; set; }
    }
}
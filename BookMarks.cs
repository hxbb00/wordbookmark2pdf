using System.Collections.Concurrent;
using System.Collections.Generic;

namespace wordbookmark2pdf
{
    /// <summary>
    /// 书签集合
    /// </summary>
    public class BookMarks
    {
        private readonly IDictionary<string, BookMark> _marks
            = new ConcurrentDictionary<string, BookMark>();
        /// <summary>
        /// 书签对应的模板名称
        /// </summary>
        public string TemplName { get; private set; }

        /// <summary>
        /// 书签项集合
        /// </summary>
        public IDictionary<string, BookMark> Marks
        {
            get
            {
                return _marks;
            }
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="templName">模板名称</param>
        public BookMarks(string templName)
        {
            TemplName = templName;
        }
    }
}
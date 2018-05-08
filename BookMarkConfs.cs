using System.Collections.Concurrent;
using System.Collections.Generic;

namespace wordbookmark2pdf
{
    /// <summary>
    /// 书签配置文件对象
    /// </summary>
    public class BookMarkConfs
    {
        private readonly IDictionary<string, BookMarks> _confs
            = new ConcurrentDictionary<string, BookMarks>();

        /// <summary>
        /// 书签集合
        /// </summary>
        public IDictionary<string, BookMarks> Confs
        {
            get
            {
                return _confs;
            }
        }
    }
}
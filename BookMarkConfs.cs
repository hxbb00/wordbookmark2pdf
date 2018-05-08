using System.Collections.Concurrent;
using System.Collections.Generic;

namespace wordbookmark2pdf
{
    /// <summary>
    /// ��ǩ�����ļ�����
    /// </summary>
    public class BookMarkConfs
    {
        private readonly IDictionary<string, BookMarks> _confs
            = new ConcurrentDictionary<string, BookMarks>();

        /// <summary>
        /// ��ǩ����
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
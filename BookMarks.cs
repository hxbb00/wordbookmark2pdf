using System.Collections.Concurrent;
using System.Collections.Generic;

namespace wordbookmark2pdf
{
    /// <summary>
    /// ��ǩ����
    /// </summary>
    public class BookMarks
    {
        private readonly IDictionary<string, BookMark> _marks
            = new ConcurrentDictionary<string, BookMark>();
        /// <summary>
        /// ��ǩ��Ӧ��ģ������
        /// </summary>
        public string TemplName { get; private set; }

        /// <summary>
        /// ��ǩ���
        /// </summary>
        public IDictionary<string, BookMark> Marks
        {
            get
            {
                return _marks;
            }
        }

        /// <summary>
        /// ���캯��
        /// </summary>
        /// <param name="templName">ģ������</param>
        public BookMarks(string templName)
        {
            TemplName = templName;
        }
    }
}
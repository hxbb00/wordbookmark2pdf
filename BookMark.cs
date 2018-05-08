namespace wordbookmark2pdf
{
    /// <summary>
    /// ��ǩ����
    /// </summary>
    public class BookMark
    {
        /// <summary>
        /// ���캯��
        /// </summary>
        /// <param name="type">��ǩ����</param>
        public BookMark(string type)
        {
            Type = type;
        }

        /// <summary>
        /// ��ǩ����
        /// </summary>
        public string Type { get; private set; }
        /// <summary>
        /// ��ǩ�ı�
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// ��ǩsql
        /// </summary>
        public string Sql { get; set; }
        /// <summary>
        /// ��ǩͼƬ
        /// </summary>
        public string PicUri { get; set; }
    }
}
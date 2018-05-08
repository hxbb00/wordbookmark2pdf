using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace wordbookmark2pdf.handlers
{
    /// <summary>
    /// ģ�崦����
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IWordBookmarkHandler<T>
    {
        /// <summary>
        /// ģ������
        /// </summary>
        string TemplateName { get; }
        /// <summary>
        /// �Զ���������
        /// </summary>
        /// <param name="t">������</param>
        /// <param name="ctx">������</param>
        /// <param name="logger">��־��¼��</param>
        /// <returns></returns>
        bool CustomTable(Table t, ReplacerContext<T> ctx, Action<string> logger = null);
        /// <summary>
        /// �Զ���ͼƬ������
        /// </summary>
        /// <param name="t">ͼƬ</param>
        /// <param name="ctx">������</param>
        /// <param name="logger">��־��¼��</param>
        /// <returns></returns>
        bool CustomPicture(DocumentFormat.OpenXml.Packaging.ImagePart t, ReplacerContext<T> ctx, Action<string> logger = null);
        /// <summary>
        /// �Զ����ı�������
        /// </summary>
        /// <param name="t">�ı�</param>
        /// <param name="ctx">������</param>
        /// <param name="logger">��־��¼��</param>
        /// <returns></returns>
        bool CustomText(Text t, ReplacerContext<T> ctx, Action<string> logger = null);
    }
}
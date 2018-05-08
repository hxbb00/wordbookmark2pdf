using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace wordbookmark2pdf.handlers
{
    /// <summary>
    /// 模板处理器
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IWordBookmarkHandler<T>
    {
        /// <summary>
        /// 模板名称
        /// </summary>
        string TemplateName { get; }
        /// <summary>
        /// 自定义表格处理器
        /// </summary>
        /// <param name="t">表格对象</param>
        /// <param name="ctx">上下文</param>
        /// <param name="logger">日志记录器</param>
        /// <returns></returns>
        bool CustomTable(Table t, ReplacerContext<T> ctx, Action<string> logger = null);
        /// <summary>
        /// 自定义图片处理器
        /// </summary>
        /// <param name="t">图片</param>
        /// <param name="ctx">上下文</param>
        /// <param name="logger">日志记录器</param>
        /// <returns></returns>
        bool CustomPicture(DocumentFormat.OpenXml.Packaging.ImagePart t, ReplacerContext<T> ctx, Action<string> logger = null);
        /// <summary>
        /// 自定义文本处理器
        /// </summary>
        /// <param name="t">文本</param>
        /// <param name="ctx">上下文</param>
        /// <param name="logger">日志记录器</param>
        /// <returns></returns>
        bool CustomText(Text t, ReplacerContext<T> ctx, Action<string> logger = null);
    }
}
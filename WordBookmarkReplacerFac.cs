using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using wordbookmark2pdf.handlers;

namespace wordbookmark2pdf
{
    /// <summary>
    /// 书签宏替换器工厂
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class WordBookmarkReplacerFac<T>
    {
        private readonly IDictionary<string, WordBookmarkReplacer<T>> _bookmarkReplacers
            = new ConcurrentDictionary<string, WordBookmarkReplacer<T>>();

        private void Register(Assembly assembly)
        {
            var types = assembly.GetTypes()
                .Where(t => null != t.GetInterface(typeof(IWordBookmarkHandler<T>).FullName) && !t.IsAbstract);

            foreach (var type in types)
            {
                var handler = Activator.CreateInstance(type) as IWordBookmarkHandler<T>;
                if (handler != null)
                {
                    var upper = handler.TemplateName.ToUpper();
                    _bookmarkReplacers.Add(upper, new WordBookmarkReplacer<T>(handler));
                }
            }

            _bookmarkReplacers.Add(DefaultWordBookmarkRepositoryHandler<T>.CustKey.ToUpper(), new WordBookmarkReplacer<T>(new DefaultWordBookmarkRepositoryHandler<T>()));
        }

        /// <summary>
        /// 执行宏替换操作
        /// </summary>
        /// <param name="ctx">上下文</param>
        /// <param name="progress">进度条</param>
        /// <param name="logger">日志</param>
        /// <param name="missMatch">匹配失败</param>
        public void Handle(ReplacerContext<T> ctx,
            ProgressChangedEventHandler progress,
            Action<string> logger, Action<string> missMatch)
        {
            progress?.Invoke(this, new ProgressChangedEventArgs(0, "开始处理"));
            if (File.Exists(ctx.SrcFilePath))
            {
                if (ctx.SrcFilePath.EndsWith(".DOCX", StringComparison.OrdinalIgnoreCase))
                {
                    var withoutExtension = ctx.SrcFileNameWithoutExt;
                    Register(ctx.GetType().Assembly);
                    if (_bookmarkReplacers.ContainsKey(withoutExtension.ToUpper()))
                    {
                        logger?.Invoke($"输入文件[{withoutExtension}]找到匹配的模板处理器");

                        _bookmarkReplacers[withoutExtension.ToUpper()].Handle(ctx, progress, logger, missMatch);
                    }
                    else
                    {
                        var custKey = DefaultWordBookmarkRepositoryHandler<T>.CustKey.ToUpper();
                        if (_bookmarkReplacers.ContainsKey(custKey))
                        {
                            logger?.Invoke($"输入文件{withoutExtension}没有找到匹配的模板处理器,将采用默认模板处理器");

                            _bookmarkReplacers[custKey].Handle(ctx, progress, logger, missMatch);
                        }
                        else
                        {
                            logger?.Invoke($"没有默认模板处理器,输入文件{withoutExtension}将不处理");
                        }
                    }
                }
                else
                {
                    logger?.Invoke($"{ctx.SrcFilePath}文件格式非法,只支持*.docx格式文件");
                }
            }
            else
            {
                logger?.Invoke($"输入文件{ctx.SrcFilePath}不存在");
            }

            progress?.Invoke(this, new ProgressChangedEventArgs(100, "处理完毕"));
        }
    }
}
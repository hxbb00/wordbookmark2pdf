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
    /// ��ǩ���滻������
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
        /// ִ�к��滻����
        /// </summary>
        /// <param name="ctx">������</param>
        /// <param name="progress">������</param>
        /// <param name="logger">��־</param>
        /// <param name="missMatch">ƥ��ʧ��</param>
        public void Handle(ReplacerContext<T> ctx,
            ProgressChangedEventHandler progress,
            Action<string> logger, Action<string> missMatch)
        {
            progress?.Invoke(this, new ProgressChangedEventArgs(0, "��ʼ����"));
            if (File.Exists(ctx.SrcFilePath))
            {
                if (ctx.SrcFilePath.EndsWith(".DOCX", StringComparison.OrdinalIgnoreCase))
                {
                    var withoutExtension = ctx.SrcFileNameWithoutExt;
                    Register(ctx.GetType().Assembly);
                    if (_bookmarkReplacers.ContainsKey(withoutExtension.ToUpper()))
                    {
                        logger?.Invoke($"�����ļ�[{withoutExtension}]�ҵ�ƥ���ģ�崦����");

                        _bookmarkReplacers[withoutExtension.ToUpper()].Handle(ctx, progress, logger, missMatch);
                    }
                    else
                    {
                        var custKey = DefaultWordBookmarkRepositoryHandler<T>.CustKey.ToUpper();
                        if (_bookmarkReplacers.ContainsKey(custKey))
                        {
                            logger?.Invoke($"�����ļ�{withoutExtension}û���ҵ�ƥ���ģ�崦����,������Ĭ��ģ�崦����");

                            _bookmarkReplacers[custKey].Handle(ctx, progress, logger, missMatch);
                        }
                        else
                        {
                            logger?.Invoke($"û��Ĭ��ģ�崦����,�����ļ�{withoutExtension}��������");
                        }
                    }
                }
                else
                {
                    logger?.Invoke($"{ctx.SrcFilePath}�ļ���ʽ�Ƿ�,ֻ֧��*.docx��ʽ�ļ�");
                }
            }
            else
            {
                logger?.Invoke($"�����ļ�{ctx.SrcFilePath}������");
            }

            progress?.Invoke(this, new ProgressChangedEventArgs(100, "�������"));
        }
    }
}
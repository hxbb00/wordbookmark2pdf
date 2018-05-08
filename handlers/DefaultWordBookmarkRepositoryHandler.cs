namespace wordbookmark2pdf.handlers
{
    class DefaultWordBookmarkRepositoryHandler<T> : WordBookmarkRepositoryHandler<T>
    {
        public override string TemplateName { get { return CustKey; } }
        public const string CustKey = "?default?";
    }
}
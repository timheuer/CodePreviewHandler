namespace TimHeuer.PreviewHandlers
{
    public class CppFormat : Manoli.Utils.CSharpFormat.CSharpFormat
    {
        protected override string Keywords
        {
            get 
            {
                string str = base.Keywords + "auto static_cast reinterpret_cast dynamic_cast safe_cast nullptr";
                return str;
            }
        }

        protected override string Preprocessors
        {
            get
            {
                return "#if #else #elif #endif #define #undef #warning "
                    + "#error #line #region #endregion #pragma #include";
            }
        }
    }
}

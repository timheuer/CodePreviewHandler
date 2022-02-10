
namespace TimHeuer.PreviewHandlers
{
    public class CSharpFormat : Manoli.Utils.CSharpFormat.CSharpFormat
    {
        protected override string Keywords
        {
            get
            {
                string str = base.Keywords + " await async";
                return str;
            }
        }
    }
}

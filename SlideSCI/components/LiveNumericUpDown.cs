using System;
using System.Globalization;
using System.Windows.Forms;

namespace SlideSCI
{
    /// <summary>输入有效数字时立即通知预览，结束编辑后再按 DecimalPlaces 格式化。</summary>
    internal sealed class LiveNumericUpDown : NumericUpDown
    {
        private bool synchronizingValue;
        private bool formattingText;

        protected override void OnTextChanged(EventArgs e)
        {
            if (!formattingText && !synchronizingValue && UserEdit &&
                decimal.TryParse(Text, NumberStyles.Number, CultureInfo.CurrentCulture, out decimal value) &&
                value >= Minimum && value <= Maximum)
            {
                synchronizingValue = true;
                // ValueChanged 中读取 Value 时不重复验证或改写当前输入。
                UserEdit = false;
                try { Value = value; }
                finally
                {
                    UserEdit = true;
                    synchronizingValue = false;
                }
            }
            // 清空、未完成的小数或超范围输入留给控件在结束编辑时验证。
            base.OnTextChanged(e);
        }

        protected override void UpdateEditText()
        {
            if (synchronizingValue) return;
            formattingText = true;
            try { base.UpdateEditText(); }
            finally { formattingText = false; }
        }
    }
}

using System;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json;

namespace SlideSCI
{
    /// <summary>FOV 始终指未裁剪原图的物理宽高，与幻灯片上的尺寸无关。</summary>
    internal sealed class ImageFieldOfView
    {
        internal const string AlternativeTextPrefix = "FOV（SlideSCI）：";
        public double Width { get; set; }
        public double Height { get; set; }
        private string unit = "μm";
        public string Unit { get => unit; set => unit = NormalizeUnit(value); }
        public string Source { get; set; } = "手动设置";

        [JsonIgnore]
        public double WidthMicrometers => Width * UnitFactor(Unit);
        [JsonIgnore]
        public double HeightMicrometers => Height * UnitFactor(Unit);
        [JsonIgnore]
        public bool IsValid => Positive(UnitFactor(Unit)) &&
            (Width == 0 || Positive(WidthMicrometers)) && (Height == 0 || Positive(HeightMicrometers)) &&
            (Positive(WidthMicrometers) || Positive(HeightMicrometers));

        /// <summary>未填写的方向保留为 0，不根据图片宽高比假定另一方向的标定。</summary>
        internal bool IsValidFor(bool vertical) => IsValid && Positive(vertical ? HeightMicrometers : WidthMicrometers);

        internal static bool Positive(double value) => value > 0 && !double.IsInfinity(value) && !double.IsNaN(value);

        internal static string NormalizeUnit(string value)
        {
            double factor = UnitFactor(value);
            if (factor == 0.001) return "nm";
            if (factor == 1) return "μm";
            if (factor == 1000) return "mm";
            if (factor == 10000) return "cm";
            if (factor == 1000000) return "m";
            if (factor == 25400) return "in";
            return value;
        }

        internal static double UnitFactor(string unit)
        {
            switch ((unit ?? "").Trim().Replace('µ', 'μ').ToLowerInvariant())
            {
                case "nm": case "nanometer": case "nanometers": return 0.001;
                case "um": case "μm": case "micron": case "microns":
                case "micrometer": case "micrometers": case "micrometre": case "micrometres": return 1;
                case "mm": case "millimeter": case "millimeters": return 1000;
                case "cm": case "centimeter": case "centimeters": return 10000;
                case "m": case "meter": case "meters": return 1000000;
                case "in": case "inch": case "inches": return 25400;
                default: return double.NaN;
            }
        }

        internal static ImageFieldOfView Read(string alternativeText)
        {
            string line = (alternativeText ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .LastOrDefault(item => item.StartsWith(AlternativeTextPrefix, StringComparison.Ordinal));
            if (line == null) return null;
            try
            {
                var fov = JsonConvert.DeserializeObject<ImageFieldOfView>(line.Substring(AlternativeTextPrefix.Length));
                return fov != null && fov.IsValid ? fov : null;
            }
            catch (JsonException) { return null; }
        }

        internal string Write(string alternativeText)
        {
            if (!IsValid) throw new InvalidOperationException("请填写至少一个大于 0 的 FOV 尺寸，且单位必须是物理长度单位。");
            // 只替换自己的记录，保留原图路径及用户已有的替换文字。
            string text = string.Join(Environment.NewLine, (alternativeText ?? "")
                .Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')
                .Where(line => !line.StartsWith(AlternativeTextPrefix, StringComparison.Ordinal))).TrimEnd('\r', '\n');
            return text + (text.Length == 0 ? "" : Environment.NewLine) +
                AlternativeTextPrefix + JsonConvert.SerializeObject(this);
        }

        public override string ToString()
        {
            if (!Positive(HeightMicrometers)) return "宽度 " + Width.ToString("0.##", CultureInfo.CurrentCulture) + " " + Unit;
            if (!Positive(WidthMicrometers)) return "高度 " + Height.ToString("0.##", CultureInfo.CurrentCulture) + " " + Unit;
            return Width.ToString("0.##", CultureInfo.CurrentCulture) + " × " +
                Height.ToString("0.##", CultureInfo.CurrentCulture) + " " + Unit;
        }
    }
}

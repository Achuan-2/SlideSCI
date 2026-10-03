using System;
using System.Drawing;
using System.Globalization;
using Newtonsoft.Json;

namespace SlideSCI
{
    internal enum ScaleBarCorner { TopLeft, BottomLeft, TopRight, BottomRight }

    internal sealed class ScaleBarSettings
    {
        public bool Vertical { get; set; }
        public double WidthMicrometers { get; set; } = 50;
        public double HeightMicrometers { get; set; } = 50;
        private string lengthUnit = "μm";
        public string LengthUnit { get => lengthUnit; set => lengthUnit = ImageFieldOfView.NormalizeUnit(value); }
        public float ThicknessPoints { get; set; } = 3;
        public bool ShowText { get; set; } = true;
        public int BarColorArgb { get; set; } = Color.White.ToArgb();
        public int FontColorArgb { get; set; } = Color.White.ToArgb();
        public float FontSize { get; set; } = 14;
        public bool Bold { get; set; } = true;
        public ScaleBarCorner Corner { get; set; } = ScaleBarCorner.BottomRight;
        public bool GroupWithImage { get; set; } = true;
        [JsonIgnore]
        public double LengthMicrometers => Vertical ? HeightMicrometers : WidthMicrometers;
        [JsonIgnore]
        public string Label => (LengthMicrometers / ImageFieldOfView.UnitFactor(LengthUnit))
            .ToString("0.##", CultureInfo.CurrentCulture) + LengthUnit;
        [JsonIgnore]
        public bool IsValid => (WidthMicrometers == 0 || ImageFieldOfView.Positive(WidthMicrometers)) &&
            (HeightMicrometers == 0 || ImageFieldOfView.Positive(HeightMicrometers)) && ImageFieldOfView.Positive(ThicknessPoints) &&
            ImageFieldOfView.Positive(FontSize) && FontSize <= 400 &&
            ImageFieldOfView.Positive(ImageFieldOfView.UnitFactor(LengthUnit)) && Enum.IsDefined(typeof(ScaleBarCorner), Corner);
        internal ScaleBarSettings Copy() => (ScaleBarSettings)MemberwiseClone();

        /// <summary>放大图的比例尺最多占当前方向 FOV 的 90%；0 表示隐藏，保持原图设置不变。</summary>
        internal ScaleBarSettings ForZoom(SizeF visibleFov, double? requestedLength = null, bool? showText = null)
        {
            var settings = Copy();
            if (showText.HasValue) settings.ShowText = showText.Value;
            double length = requestedLength ?? LengthMicrometers;
            if (length != 0)
            {
                double fieldLength = Vertical ? visibleFov.Height : visibleFov.Width;
                if (!ImageFieldOfView.Positive(length) || !ImageFieldOfView.Positive(fieldLength))
                    throw new InvalidOperationException(Vertical ? "纵向比例尺需要 FOV 高度。" : "横向比例尺需要 FOV 宽度。");
                length = Math.Min(length, fieldLength * 0.9);
            }
            if (Vertical) settings.HeightMicrometers = length;
            else settings.WidthMicrometers = length;
            return settings;
        }
        internal static ScaleBarSettings Parse(string json)
        {
            try
            {
                var settings = JsonConvert.DeserializeObject<ScaleBarSettings>(json ?? "");
                return settings != null && settings.IsValid ? settings : null;
            }
            catch (JsonException) { return null; }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace SlideSCI
{
    /// <summary>仅读取 TIFF 首个 IFD 的标定信息，不解码图像或加载整个时间序列。</summary>
    internal static class TiffFieldOfViewReader
    {
        internal const string ResolutionSource = "TIFF 分辨率（可能为打印 DPI，请核对）";
        private const int MaximumMetadataBytes = 4 * 1024 * 1024;

        internal static ImageFieldOfView TryRead(string path)
        {
            if (string.IsNullOrEmpty(path) ||
                !(string.Equals(Path.GetExtension(path), ".tif", StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(Path.GetExtension(path), ".tiff", StringComparison.OrdinalIgnoreCase))) return null;
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new DirectoryReader(stream))
                {
                    Dictionary<int, byte[]> tags = reader.ReadTags();
                    double width = reader.Scalar(tags, 256);
                    double height = reader.Scalar(tags, 257);
                    if (!ImageFieldOfView.Positive(width) || !ImageFieldOfView.Positive(height)) return null;
                    string description = tags.TryGetValue(270, out byte[] bytes)
                        ? Encoding.UTF8.GetString(bytes).TrimEnd('\0') : "";
                    ImageFieldOfView ome = ReadOme(description, width, height);
                    if (ome != null) return ome;

                    double xResolution = reader.Scalar(tags, 282);
                    double yResolution = reader.Scalar(tags, 283);
                    var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (string line in description.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        int equals = line.IndexOf('=');
                        if (equals > 0) values[line.Substring(0, equals).Trim()] = line.Substring(equals + 1).Trim();
                    }
                    if (values.ContainsKey("ImageJ"))
                    {
                        // ImageJ 的 X/YResolution 为每个标定单位内的像素数；spacing 是 Z 轴间距，不能当成 XY 标定。
                        if (!values.TryGetValue("unit", out string unit) ||
                            double.IsNaN(ImageFieldOfView.UnitFactor(unit))) return null;
                        double pixelWidth = Number(values, "pixel_width");
                        double pixelHeight = Number(values, "pixel_height");
                        if (!ImageFieldOfView.Positive(pixelWidth)) pixelWidth = 1 / xResolution;
                        if (!ImageFieldOfView.Positive(pixelHeight)) pixelHeight = 1 / yResolution;
                        return Create(width * pixelWidth, height * pixelHeight, unit, "ImageJ TIFF 标定");
                    }
                    // 普通 TIFF 的分辨率也可能只是打印 DPI；记录来源，设置窗口提示核对。
                    double resolutionUnit = reader.Scalar(tags, 296);
                    string standardUnit = resolutionUnit == 2 ? "in" : resolutionUnit == 3 ? "cm" : null;
                    // 标签单位用于物理换算；显微图像的 FOV 以 μm 展示，避免 cm/in
                    // 在两位小数的编辑窗口中变成 0.03、0.00，掩盖真实标定。
                    double micrometersPerUnit = ImageFieldOfView.UnitFactor(standardUnit);
                    return standardUnit == null ? null : Create(width / xResolution * micrometersPerUnit,
                        height / yResolution * micrometersPerUnit, "μm", ResolutionSource);
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException ||
                ex is InvalidDataException || ex is OverflowException || ex is XmlException || ex is ArgumentException)
            {
                Debug.WriteLine($"无法读取 TIFF FOV：{path}：{ex.Message}");
                return null;
            }
        }

        private static ImageFieldOfView ReadOme(string description, double width, double height)
        {
            if (string.IsNullOrWhiteSpace(description) || !description.TrimStart().StartsWith("<")) return null;
            try
            {
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                using (var xml = XmlReader.Create(new StringReader(description), settings))
                {
                    XDocument document = XDocument.Load(xml);
                    if (document.Root?.Name.LocalName != "OME") return null;
                    // 多 series 的 TIFF 只对应插入的首张画面，不使用其他 series 的标定。
                    XElement pixels = document.Descendants().FirstOrDefault(item => item.Name.LocalName == "Pixels");
                    if (pixels == null || Parse((string)pixels.Attribute("SizeX")) != width ||
                        Parse((string)pixels.Attribute("SizeY")) != height) return null;
                    double x = Parse((string)pixels.Attribute("PhysicalSizeX"));
                    double y = Parse((string)pixels.Attribute("PhysicalSizeY"));
                    double xUnit = ImageFieldOfView.UnitFactor((string)pixels.Attribute("PhysicalSizeXUnit") ?? "μm");
                    double yUnit = ImageFieldOfView.UnitFactor((string)pixels.Attribute("PhysicalSizeYUnit") ?? "μm");
                    return Create(width * x * xUnit, height * y * yUnit, "μm", "OME-TIFF 像素标定");
                }
            }
            catch (XmlException) { return null; }
        }

        private static double Number(IDictionary<string, string> values, string key) =>
            values.TryGetValue(key, out string value) ? Parse(value) : double.NaN;
        private static double Parse(string value) =>
            double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : double.NaN;
        private static ImageFieldOfView Create(double width, double height, string unit, string source)
        {
            var fov = new ImageFieldOfView { Width = width, Height = height, Unit = unit, Source = source };
            return fov.IsValid ? fov : null;
        }

        /// <summary>有界读取经典 TIFF/BigTIFF，兼容大小端；只保留需要的标签。</summary>
        private sealed class DirectoryReader : IDisposable
        {
            private readonly BinaryReader reader;
            private readonly bool littleEndian;
            private readonly Dictionary<int, int> types = new Dictionary<int, int>();
            internal DirectoryReader(Stream stream)
            {
                reader = new BinaryReader(stream, Encoding.UTF8, true);
                byte first = reader.ReadByte();
                byte second = reader.ReadByte();
                if (first != second || (first != 'I' && first != 'M')) throw new InvalidDataException("不是 TIFF。");
                littleEndian = first == 'I';
            }

            internal Dictionary<int, byte[]> ReadTags()
            {
                ulong version = Unsigned(Read(2), 0, 2);
                bool big = version == 43;
                if (version != 42 && !big) throw new InvalidDataException("未知 TIFF 版本。");
                if (big && (Unsigned(Read(2), 0, 2) != 8 || Unsigned(Read(2), 0, 2) != 0))
                    throw new InvalidDataException("未知 BigTIFF 偏移类型。");
                Seek(Unsigned(Read(big ? 8 : 4), 0, big ? 8 : 4));
                ulong count = Unsigned(Read(big ? 8 : 2), 0, big ? 8 : 2);
                if (count > 4096) throw new InvalidDataException("TIFF 标签过多。");
                var tags = new Dictionary<int, byte[]>();
                for (ulong i = 0; i < count; i++)
                {
                    byte[] entry = Read(big ? 20 : 12);
                    int tag = (int)Unsigned(entry, 0, 2);
                    if (tag != 256 && tag != 257 && tag != 270 && tag != 282 && tag != 283 && tag != 296) continue;
                    int type = (int)Unsigned(entry, 2, 2);
                    int size = type == 2 ? 1 : type == 3 ? 2 : type == 4 ? 4 : type == 5 || type == 16 ? 8 : 0;
                    ulong items = Unsigned(entry, 4, big ? 8 : 4);
                    if (size == 0 || items == 0 || items > (ulong)(MaximumMetadataBytes / size)) continue;
                    int length = checked((int)items * size);
                    int valueOffset = big ? 12 : 8;
                    byte[] data;
                    if (length <= (big ? 8 : 4)) data = entry.Skip(valueOffset).Take(length).ToArray();
                    else
                    {
                        long position = reader.BaseStream.Position;
                        Seek(Unsigned(entry, valueOffset, big ? 8 : 4));
                        data = Read(length);
                        reader.BaseStream.Position = position;
                    }
                    tags[tag] = data;
                    types[tag] = type;
                }
                return tags;
            }

            internal double Scalar(IDictionary<int, byte[]> tags, int tag)
            {
                if (!tags.TryGetValue(tag, out byte[] data)) return double.NaN;
                int type = types[tag];
                if (type == 5)
                {
                    ulong denominator = Unsigned(data, 4, 4);
                    return denominator == 0 ? double.NaN : (double)Unsigned(data, 0, 4) / denominator;
                }
                return type == 3 || type == 4 || type == 16
                    ? Unsigned(data, 0, type == 3 ? 2 : type == 4 ? 4 : 8) : double.NaN;
            }

            private ulong Unsigned(byte[] data, int offset, int length)
            {
                ulong value = 0;
                for (int i = 0; i < length; i++) value = (value << 8) | data[offset + (littleEndian ? length - i - 1 : i)];
                return value;
            }
            private byte[] Read(int length)
            {
                byte[] data = reader.ReadBytes(length);
                if (data.Length != length) throw new EndOfStreamException();
                return data;
            }
            private void Seek(ulong position)
            {
                if (position > (ulong)reader.BaseStream.Length) throw new InvalidDataException("TIFF 偏移越界。");
                reader.BaseStream.Position = (long)position;
            }
            public void Dispose() => reader.Dispose();
        }
    }
}

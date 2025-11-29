using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using Markdig;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Font = System.Drawing.Font;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private PowerPoint.Application app;
        private float copiedWidth;
        private float copiedHeight;

        private List<float> copiedLeft = new List<float>();
        private List<float> copiedTop = new List<float>();

        private float cropLeft;
        private float cropRight;
        private float cropTop;
        private float cropBottom;
        private bool hasCopiedCrop = false;
        private float originalHeight; // 添加变量存储原始图片高度
        private float currentCropedHeight;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;

            // Load Image Title Settings
            fontNameEditBox.Text = Properties.Settings.Default.TitleFontName;
            fontSizeEditBox.Text = Properties.Settings.Default.TitleFontSize;
            distanceFromBottomEditBox.Text = Properties.Settings.Default.TitleDistanceFromBottom;
            titleTextEditBox.Text = Properties.Settings.Default.TitleText;
            autoGroupCheckBox.Checked = Properties.Settings.Default.AutoGroup;

            // Load Image Label Settings
            labelOffsetXEditBox.Text = Properties.Settings.Default.LabelOffsetX;
            labelOffsetYEditBox.Text = Properties.Settings.Default.LabelOffsetY;
            labelTemplateComboBox.Text = Properties.Settings.Default.LabelTemplate;
            labelFontNameEditBox.Text = Properties.Settings.Default.LabelFontName;
            labelFontSizeEditBox.Text = Properties.Settings.Default.LabelFontSize;
            labelBoldcheckBox.Checked = Properties.Settings.Default.LabelBold;
            labelIndex.Text = "1";
            labelIndexUpdatecheckBox.Checked = true;

            // Load Image Auto Align Settings
            imgAutoAlignSortTypeDropDown.SelectedItemIndex = Properties
                .Settings
                .Default
                .imgAutoAlignSortType;
            imgAutoAlign_colNum.Text = Properties.Settings.Default.ColNum;
            imgAutoAlign_colSpace.Text = Properties.Settings.Default.ColSpace;
            imgAutoAlign_rowSpace.Text = Properties.Settings.Default.RowSpace;
            imgWidthEditBpx.Text = Properties.Settings.Default.ImgWidth;
            imgHeightEditBox.Text = Properties.Settings.Default.ImgHeight;
            imgAutoAlignAlignTypeDropDown.SelectedItemIndex = Properties
                .Settings
                .Default
                .imgAutoAlignAlignType;
            excludeTextcheckBox.Checked = Properties.Settings.Default.imgAutoAlighExcludeText;
            excludeTextcheckBox2.Checked = Properties.Settings.Default.imgAddTitleExcludeText;
            titleCenterCheckbox.Checked = Properties.Settings.Default.imgAddTitleCenter;
            // insertMarkdown
            toggleBackgroundCheckBox.Checked = Properties.Settings.Default.ToggleBackground;

            // Add event handlers for text changed events
            fontNameEditBox.TextChanged += SaveSettings;
            fontSizeEditBox.TextChanged += SaveSettings;
            distanceFromBottomEditBox.TextChanged += SaveSettings;
            titleTextEditBox.TextChanged += SaveSettings;
            autoGroupCheckBox.Click += SaveSettings;

            labelOffsetXEditBox.TextChanged += SaveSettings;
            labelOffsetYEditBox.TextChanged += SaveSettings;
            labelTemplateComboBox.TextChanged += SaveSettings;
            labelFontNameEditBox.TextChanged += SaveSettings;
            labelFontSizeEditBox.TextChanged += SaveSettings;

            imgAutoAlignSortTypeDropDown.SelectionChanged += SaveSettings;
            imgAutoAlign_colNum.TextChanged += SaveSettings;
            imgAutoAlign_colSpace.TextChanged += SaveSettings;
            imgAutoAlign_rowSpace.TextChanged += SaveSettings;
            imgWidthEditBpx.TextChanged += SaveSettings;
            imgHeightEditBox.TextChanged += SaveSettings;
            imgAutoAlignAlignTypeDropDown.SelectionChanged += SaveSettings;
            excludeTextcheckBox.Click += SaveSettings;
            excludeTextcheckBox2.Click += SaveSettings;
            titleCenterCheckbox.Click += SaveSettings;
            labelBoldcheckBox.Click += SaveSettings;

            toggleBackgroundCheckBox.Click += SaveSettings;
            // exportImageButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportImageButton_Click); // Already set in Designer.cs
            iniCombobox();
        }

        /// <summary>
        /// 初始化下拉框的值
        /// </summary>
        public void iniCombobox()
        {
            //字体名
            List<string> FontNames = new List<string>()
            {
                "Arial",
                "Times new Roman",
                "微软雅黑",
                "黑体",
                "方正兰亭黑体",
                "仿宋",
                "楷体",
                "宋体",
                "新宋体",
                "华文中宋",
                "华文仿宋",
                "华文行楷",
                "华文新魏",
                "汉仪综艺体简",
                "思源黑体",
                "思源宋体",
                "庞门正道标题体",
                "方正清刻本悦宋",
                "文悦新青年体",
                "演示新手书",
            };
            FreshCombobox(fontNameEditBox, FontNames);
            FreshCombobox(labelFontNameEditBox, FontNames);
            //字号
            List<string> FontSizes = new List<string>()
            {
                "2",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "10",
                "11",
                "12",
                "13",
                "14",
                "15",
                "16",
                "18",
                "20",
                "22",
                "24",
                "26",
                "28",
                "30",
                "40",
                "50",
                "60",
                "80",
                "100",
                "120",
                "150",
                "200",
            };
            FreshCombobox(fontSizeEditBox, FontSizes);
            FreshCombobox(labelFontSizeEditBox, FontSizes);
            //图片宽度和高度
            List<string> PicSizes = new List<string>()
            {
                "0cm",
                "0.5cm",
                "1cm",
                "2cm",
                "3cm",
                "4cm",
                "5cm",
                "6cm",
                "7cm",
                "8cm",
                "9cm",
                "10cm",
                "12cm",
                "15cm",
                "20cm",
                "25cm",
                "30cm",
                "35cm",
                "40cm",
                "45cm",
                "50cm",
                "60cm",
                "70cm",
                "80cm",
                "100cm",
                "120cm",
                "150cm",
                "200cm",
            };
            FreshCombobox(imgWidthEditBpx, PicSizes);
            FreshCombobox(imgHeightEditBox, PicSizes);
            //图下距离
            List<string> PicDistance = new List<string>()
            {
                "0",
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "10",
                "11",
                "12",
                "13",
                "14",
                "15",
                "20",
                "25",
                "30",
                "35",
                "40",
                "45",
                "50",
                "55",
                "60",
                "65",
                "70",
                "75",
                "80",
                "90",
                "100",
                "120",
                "150",
                "200",
                "500",
            };
            FreshCombobox(distanceFromBottomEditBox, PicDistance);
            //XY偏移
            List<string> OffsetValues = new List<string>()
            {
                "-40",
                "-30",
                "-20",
                "-10",
                "-15",
                "-10",
                "-9",
                "-8",
                "-7",
                "-6",
                "-5",
                "-4",
                "-3",
                "-2",
                "-1",
                "0",
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "10",
                "15",
                "20",
                "25",
                "30",
                "40",
            };
            FreshCombobox(labelOffsetYEditBox, OffsetValues);
            FreshCombobox(labelOffsetXEditBox, OffsetValues);
            //列间距
            List<string> columnGap = new List<string>()
            {
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "10",
                "11",
                "12",
                "13",
                "14",
                "15",
                "16",
                "17",
                "18",
                "19",
                "20",
                "21",
                "22",
                "23",
                "24",
                "25",
                "30",
                "35",
                "40",
                "45",
                "50",
                "55",
                "60",
                "80",
            };
            FreshCombobox(imgAutoAlign_colSpace, columnGap);
            //行间距
            List<string> RowGap = new List<string>()
            {
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "10",
                "11",
                "12",
                "13",
                "14",
                "15",
                "16",
                "17(≈08字框高)",
                "18(≈09字框高)",
                "20(≈10字框高)",
                "22(≈12字框高)",
                "25(≈14字框高)",
                "27(≈16字框高)",
                "29(≈18字框高)",
                "32(≈20字框高)",
                "34(≈22字框高)",
                "37(≈24字框高)",
                "39(≈26字框高)",
                "41(≈28字框高)",
                "44(≈30字框高)",
                "51(≈35字框高)",
                "56(≈40字框高)",
                "68(≈50字框高)",
                "80(≈60字框高)",
            };
            FreshCombobox(imgAutoAlign_rowSpace, RowGap);
            //列数量
            List<string> columNums = new List<string>()
            {
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "10",
                "11",
                "12",
                "13",
                "14",
                "15",
                "16",
                "17",
                "18",
                "19",
                "20",
            };
            FreshCombobox(imgAutoAlign_colNum, columNums);
        }

        /// <summary>
        /// RibbonComboBox下拉值初始化
        /// </summary>
        /// <param name="BOX"></param>
        private void FreshCombobox(RibbonComboBox BOX, List<string> itemLabel)
        {
            BOX.Items.Clear();
            // 使用 LINQ 创建 RibbonDropDownItem 并添加到 RibbonDropDown 中
            itemLabel
                .Select(x =>
                {
                    RibbonDropDownItem item = Globals
                        .Factory.GetRibbonFactory()
                        .CreateRibbonDropDownItem();
                    item.Label = x; // 设置项的显示文本
                    return item;
                })
                .ToList()
                .ForEach(item => BOX.Items.Add(item)); // 将项添加到下拉菜单中
        }

        private void SaveSettings(object sender, RibbonControlEventArgs e)
        {
            // Save Image Title Settings
            Properties.Settings.Default.TitleFontName = fontNameEditBox.Text;
            Properties.Settings.Default.TitleFontSize = fontSizeEditBox.Text;
            Properties.Settings.Default.TitleDistanceFromBottom = distanceFromBottomEditBox.Text;
            Properties.Settings.Default.TitleText = titleTextEditBox.Text;
            Properties.Settings.Default.AutoGroup = autoGroupCheckBox.Checked;

            // Save Image Label Settings
            Properties.Settings.Default.LabelOffsetX = labelOffsetXEditBox.Text;
            Properties.Settings.Default.LabelOffsetY = labelOffsetYEditBox.Text;
            Properties.Settings.Default.LabelTemplate = labelTemplateComboBox.Text;
            Properties.Settings.Default.LabelFontName = labelFontNameEditBox.Text;
            Properties.Settings.Default.LabelFontSize = labelFontSizeEditBox.Text;
            Properties.Settings.Default.LabelBold = labelBoldcheckBox.Checked;
            // Save Image Auto Align Settings
            Properties.Settings.Default.imgAutoAlignSortType =
                imgAutoAlignSortTypeDropDown.SelectedItemIndex;
            Properties.Settings.Default.ColNum = imgAutoAlign_colNum.Text;
            Properties.Settings.Default.ColSpace = imgAutoAlign_colSpace.Text;
            Properties.Settings.Default.RowSpace = imgAutoAlign_rowSpace.Text;
            Properties.Settings.Default.ImgWidth = imgWidthEditBpx.Text;
            Properties.Settings.Default.ImgHeight = imgHeightEditBox.Text;
            Properties.Settings.Default.imgAutoAlignAlignType =
                imgAutoAlignAlignTypeDropDown.SelectedItemIndex;
            Properties.Settings.Default.imgAutoAlighExcludeText = excludeTextcheckBox.Checked;
            Properties.Settings.Default.imgAddTitleExcludeText = excludeTextcheckBox2.Checked;
            Properties.Settings.Default.imgAddTitleCenter = titleCenterCheckbox.Checked;
            // Save insertMarkdwon
            Properties.Settings.Default.ToggleBackground = toggleBackgroundCheckBox.Checked;

            // 保存导出设置 (如果将来添加UI控件进行修改)
            // Properties.Settings.Default.ExportFormat = exportFormatComboBox.Text;
            // Properties.Settings.Default.ExportDPI = int.Parse(exportDpiEditBox.Text);

            // Save all settings
            Properties.Settings.Default.Save();

            // 弹窗显示已保存
            // MessageBox.Show("设置已保存");
        }

        /// <summary>
        /// 图片加标题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddTitleToImage(object sender, RibbonControlEventArgs e)
        {
            AddTitleFun(true);
        }
        private void AddTopTitleToImage(object sender, RibbonControlEventArgs e)
        {
            AddTitleFun(false);
        }


    /// <summary>
    /// 图片加标题
    /// </summary>
    /// <param name="isBottomTitle">true为下标题，false为上标题</param>
    private void AddTitleFun(bool isBottomTitle = true)
    {
        PowerPoint.Application app = Globals.ThisAddIn.Application;
        Slide slide = app.ActiveWindow.View.Slide;
        Selection sel = app.ActiveWindow.Selection;
        bool autoGroup = autoGroupCheckBox.Checked; // 自动编组
        List<ShapeRange> allshapesName = new List<ShapeRange>(); // 需要编组的对象集合
        List<Shape> allshapes = new List<Shape>(); // 编组后的对象

        if (sel.Type == PpSelectionType.ppSelectionShapes)
        {
            float fontSize = float.Parse(fontSizeEditBox.Text); // 字号
            float distanceFromBottom = float.Parse(distanceFromBottomEditBox.Text); // 图下距离
            string fontName = fontNameEditBox.Text; // 字体名称
            string titleText = titleTextEditBox.Text; // 标题文本
            int count = 1;
            float tolerance = 10f; // 通常图片排列错位容差，10就够用
            ShapeRange sel2 = GetSortedSelection(sel, tolerance);
            var selectedImgShape = new List<Shape>();

            foreach (Shape shape in sel.ShapeRange)
            {
                Office.MsoShapeType objType = shape.Type;
                // 是否排除文本框、形状等格式。excludeTextcheckBox2.Checked，则排除
                if (
                    excludeTextcheckBox2.Checked
                    && (
                        objType is Office.MsoShapeType.msoTextBox
                        || objType is Office.MsoShapeType.msoAutoShape
                    )
                )
                {
                    continue;
                }

                // 弹窗显示objType
                //MessageBox.Show($"对象类型: {objType}");
                // 检查是否为支持的类型：图片、视频、媒体对象
                if (
                    objType == Office.MsoShapeType.msoPicture
                    || objType == Office.MsoShapeType.msoMedia
                    || objType == Office.MsoShapeType.msoLinkedPicture
                    || objType == Office.MsoShapeType.msoEmbeddedOLEObject
                    || objType == Office.MsoShapeType.msoLinkedOLEObject
                    || objType == Office.MsoShapeType.msoPlaceholder
                    || (
                        !excludeTextcheckBox2.Checked
                        && (
                            objType == Office.MsoShapeType.msoTextBox
                            || objType == Office.MsoShapeType.msoAutoShape
                        )
                    )
                )
                {
                    selectedImgShape.Add(shape);
                }
            }

            foreach (Shape selectedShape in selectedImgShape)
            {
                try
                {
                    // 根据参数决定标题位置
                    float titleTop;
                    if (isBottomTitle)
                    {
                        // 下标题：图片底部 + 距离
                        titleTop = selectedShape.Top + selectedShape.Height + distanceFromBottom;
                    }
                    else
                    {
                        // 上标题：图片顶部 - 标题高度 - 距离
                        titleTop = selectedShape.Top - (fontSize * 2) - distanceFromBottom;
                    }

                    Shape titleShape = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        selectedShape.Left,
                        titleTop,
                        selectedShape.Width,
                        fontSize * 2
                    );

                    // 设置标题文本和格式
                    titleShape.TextFrame.TextRange.Text = titleText;
                    titleShape.TextFrame.TextRange.Font.Size = fontSize;
                    titleShape.TextFrame.TextRange.Font.NameFarEast = fontName; // Ensure FarEast font is set
                    titleShape.TextFrame.TextRange.Font.Name = fontName; // Ensure font is set
                    // 标题是否居中
                    if (titleCenterCheckbox.Checked)
                    {
                        titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                    }
                    else{
                        titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

                    }


                    // 形状中的文字是否自动换行
                    titleShape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                    // 自动调整文本框大小
                    titleShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

                    // 设置文本框宽度
                    titleShape.Width = selectedShape.Width;
                    titleShape.Left = selectedShape.Left; // 设置文本框左对齐

                    allshapesName.Add(
                        slide.Shapes.Range(new string[] { selectedShape.Name, titleShape.Name })
                    );

                    // 自动选择
                    if (count == 1)
                    {
                        titleShape.Select(Office.MsoTriState.msoTrue);
                    }
                    else
                    {
                        titleShape.Select(Office.MsoTriState.msoFalse);
                    }
                    count++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"为对象 '{selectedShape.Name}' 添加标题时出错: {ex.Message}"
                    );
                    continue; // 继续处理下一个对象
                }
            }

            if (selectedImgShape.Count == 0)
            {
                MessageBox.Show("没有找到支持添加标题的对象。请选择图片、视频或其他媒体对象。");
                return;
            }
        }
        else
        {
            MessageBox.Show("请选择需要增加标题的图片、形状、视频对象.");
        }

        // 自动编组
        if (autoGroup)
        {
            foreach (var shapeRange2 in allshapesName)
            {
                Shape GroupObj;
                try
                {
                    GroupObj = shapeRange2.Group();
                    allshapes.Add(GroupObj);
                    SelectMultipleShapes(allshapes);
                }
                catch (Exception ex)
                {
                    try
                    {
                        shapeRange2.Copy();
                        shapeRange2.Delete();

                        ShapeRange pastedShapes = slide.Shapes.Paste();

                        GroupObj = pastedShapes.Group();
                        allshapes.Add(GroupObj);
                        SelectMultipleShapes(allshapes);
                    }
                    catch (Exception innerEx)
                    {
                        MessageBox.Show($"编组失败：{innerEx.Message}");
                        continue;
                    }
                }
            }
        }
    }



        /// <summary>
        /// 选择集排序
        /// </summary>
        /// <param name="initialSelection">原始选择集</param>
        /// <returns></returns>
        public ShapeRange GetSortedSelection(Selection initialSelection, float tolerance)
        {
            try
            {
                // 确保选择集中有形状对象
                if (initialSelection.ShapeRange.Count == 0)
                {
                    MessageBox.Show("初始选择集中未包含任何形状。");
                    return null;
                }

                // 将选择集中的形状转换为 List<PowerPoint.Shape>
                List<Shape> shapes = initialSelection.ShapeRange.Cast<Shape>().ToList();

                // 根据 X 从小到大、Y 从大到小排序
                var sortedShapes = shapes
                    .OrderBy(shape => shape.Top + tolerance) // Y 坐标从小到大
                    .ThenByDescending(shape => (shape.Left + tolerance) * -1) // X 坐标从大到小
                    .ToList();

                // 将排序后的形状转换为 ShapeRange
                object[] shapeNames = sortedShapes.Select(shape => (object)shape.Name).ToArray();
                ShapeRange sortedShapeRange = initialSelection
                    .ShapeRange[1]
                    .Parent.Shapes.Range(shapeNames);

                return sortedShapeRange;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"排序失败：{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 选择多个对象
        /// </summary>
        /// <param name="shapeNames"></param>
        public void SelectMultipleShapes(List<Shape> shapesToSelect)
        {
            try
            {
                // 获取当前 PowerPoint 应用实例
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

                // 检查是否处于普通视图
                if (pptApp.ActiveWindow.View.Type != PpViewType.ppViewNormal)
                {
                    MessageBox.Show("请切换到普通视图以操作形状。");
                    return;
                }

                // 获取当前幻灯片
                Slide currentSlide = pptApp.ActiveWindow.View.Slide as Slide;
                if (currentSlide == null)
                {
                    MessageBox.Show("未找到活动幻灯片。");
                    return;
                }

                // 提取形状名称列表
                List<object> shapeNames = new List<object>();
                foreach (Shape shape in shapesToSelect)
                {
                    shapeNames.Add((object)shape.Name);
                }

                // 选中所有形状
                if (shapeNames.Count > 0)
                {
                    ShapeRange selectedShapes = currentSlide.Shapes.Range(shapeNames.ToArray());
                    selectedShapes.Select();
                    pptApp.ActiveWindow.Activate(); // 确保窗口焦点
                }
                else
                {
                    MessageBox.Show("未提供有效的形状列表。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败：{ex.Message}");
            }
        }

        private void pasteImgWidthHeight_Click(object sender, RibbonControlEventArgs e)
        {
            if (copiedWidth <= 0 || copiedHeight <= 0)
            {
                MessageBox.Show("Invalid copied dimensions. Please copy the dimensions again.");
                return;
            }

            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
                {
                    shape.Width = copiedWidth;
                    shape.Height = copiedHeight;
                }
            }
            else
            {
                MessageBox.Show("Please select an image to paste dimensions.");
            }
        }

        private void copyImgWidthHeight_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape shape = sel.ShapeRange[1];
                copiedWidth = shape.Width;
                copiedHeight = shape.Height;
                // MessageBox.Show("Image dimensions copied!");
            }
            else
            {
                MessageBox.Show("Please select an image to copy dimensions.");
            }
        }

        private void copyImgWidth_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape shape = sel.ShapeRange[1];
                copiedWidth = shape.Width;
                // MessageBox.Show("Image width copied!");
            }
            else
            {
                MessageBox.Show("Please select an image to copy width.");
            }
        }

        private void pasteImgWidth_Click(object sender, RibbonControlEventArgs e)
        {
            if (copiedWidth <= 0)
            {
                MessageBox.Show("Invalid copied width. Please copy the width again.");
                return;
            }

            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
                {
                    shape.LockAspectRatio = Office.MsoTriState.msoTrue; // Lock aspect ratio
                    shape.Width = copiedWidth;
                }
            }
            else
            {
                MessageBox.Show("Please select an image to paste width.");
            }
        }

        private void copyImgHeight_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape shape = sel.ShapeRange[1];
                copiedHeight = shape.Height;
                // MessageBox.Show("Image height copied!");
            }
            else
            {
                MessageBox.Show("Please select an image to copy height.");
            }
        }

        private void pasteImgHeight_Click(object sender, RibbonControlEventArgs e)
        {
            if (copiedHeight <= 0)
            {
                MessageBox.Show("Invalid copied height. Please copy the height again.");
                return;
            }

            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
                {
                    shape.LockAspectRatio = Office.MsoTriState.msoTrue; // Lock aspect ratio
                    shape.Height = copiedHeight;
                }
            }
            else
            {
                MessageBox.Show("Please select an image to paste height.");
            }
        }

        private void copyPosition_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                copiedLeft.Clear();
                copiedTop.Clear();
                foreach (Shape shape in sel.ShapeRange)
                {
                    copiedLeft.Add(shape.Left + shape.Width / 2);
                    copiedTop.Add(shape.Top + shape.Height / 2);
                }
                // MessageBox.Show($"Copied positions of {sel.ShapeRange.Count} shapes");
            }
            else
            {
                MessageBox.Show("Please select shapes to copy positions.");
            }
        }

        private void pastePosition_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                int count = Math.Min(sel.ShapeRange.Count, copiedLeft.Count);
                if (count == 0)
                {
                    MessageBox.Show("No positions copied yet.");
                    return;
                }

                for (int i = 0; i < count; i++)
                {
                    sel.ShapeRange[i + 1].Left = copiedLeft[i] - sel.ShapeRange[i + 1].Width / 2;
                    sel.ShapeRange[i + 1].Top = copiedTop[i] - sel.ShapeRange[i + 1].Height / 2;
                }

                // if (sel.ShapeRange.Count > copiedLeft.Count)
                // {
                //     MessageBox.Show("More shapes selected than positions copied. Only the first " + copiedLeft.Count + " shapes were positioned.");
                // }
            }
            else
            {
                MessageBox.Show("Please select shapes to paste positions.");
            }
        }

        /// <summary>
        /// 图片排列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void imgAutoAlign_Click(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        /// <summary>
        /// 图片对齐排列
        /// </summary>
        private void AlignPics()
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                int colNum;
                float colSpace;
                float rowSpace;
                float customWidth = 0;
                float customHeight = 0;

                // Input validation
                if (!int.TryParse(imgAutoAlign_colNum.Text, out colNum) || colNum <= 0)
                {
                    MessageBox.Show("请输入有效的列数量。");
                    return;
                }

                if (!float.TryParse(imgAutoAlign_colSpace.Text, out colSpace) || colSpace < 0)
                {
                    MessageBox.Show("请输入有效的列间距。");
                    return;
                }

                if (
                    !float.TryParse(
                        imgAutoAlign_rowSpace.Text.Split(new char[] { '(', ' ' })[0],
                        out rowSpace
                    )
                    || rowSpace < 0
                )
                {
                    rowSpace = colSpace;
                }

                bool useCustomWidth =
                    float.TryParse(imgWidthEditBpx.Text.Split('c')[0], out customWidth)
                    && customWidth > 0;
                bool useCustomHeight =
                    float.TryParse(imgHeightEditBox.Text.Split('c')[0], out customHeight)
                    && customHeight > 0;
                customWidth = (float)(customWidth * 28.34646);
                customHeight = (float)(customHeight * 28.34646);
                var selectedImgShape = new List<Shape>();
                foreach (Shape shape in sel.ShapeRange)
                {
                    // Skip text boxes if excludeTextcheckBox is checked
                    Office.MsoShapeType objType = shape.Type;
                    if (
                        excludeTextcheckBox.Checked
                        && (
                            objType is Office.MsoShapeType.msoTextBox
                            || objType is Office.MsoShapeType.msoAutoShape
                            || objType is Office.MsoShapeType.msoMedia
                        )
                    )
                    {
                        continue;
                    }
                    selectedImgShape.Add(shape);
                }

                List<Shape> shapesToArrange = new List<Shape>();

                if (imgAutoAlignSortTypeDropDown.SelectedItemIndex == 0)
                {
                    // Create groups based on vertical position
                    var groups = new List<ImageGroup>();
                    var shapes = new List<Shape>();
                    foreach (Shape shape in selectedImgShape)
                    {
                        shapes.Add(shape);
                    }

                    // Group shapes based on vertical overlap
                    foreach (var shape in shapes)
                    {
                        bool addedToExistingGroup = false;
                        foreach (var group in groups)
                        {
                            if (group.OverlapsWith(shape))
                            {
                                group.AddShape(shape);
                                addedToExistingGroup = true;
                                break;
                            }
                        }

                        if (!addedToExistingGroup)
                        {
                            var newGroup = new ImageGroup();
                            newGroup.AddShape(shape);
                            groups.Add(newGroup);
                        }
                    }

                    // Sort shapes within each group by x position
                    foreach (var group in groups)
                    {
                        group.Shapes.Sort((a, b) => a.Left.CompareTo(b.Left));
                    }

                    // Sort groups by MinTop
                    groups.Sort((a, b) => a.MinTop.CompareTo(b.MinTop));

                    // Flatten all shapes from all groups into a single list for arrangement
                    foreach (var group in groups)
                    {
                        shapesToArrange.AddRange(group.Shapes);
                    }
                }
                else
                {
                    // Use shapes in their original order
                    foreach (Shape shape in selectedImgShape)
                    {
                        shapesToArrange.Add(shape);
                    }
                }
                // Now Align image
                float startX = shapesToArrange[0].Left;
                float startY = shapesToArrange[0].Top;
                float currentY = shapesToArrange[0].Top;

                if (imgAutoAlignAlignTypeDropDown.SelectedItemIndex == 0)
                {
                    // 1. 预先将图片分配到列
                    List<List<Shape>> columns = new List<List<Shape>>();
                    for (int i = 0; i < colNum; i++)
                    {
                        columns.Add(new List<Shape>());
                    }

                    for (int i = 0; i < shapesToArrange.Count; i++)
                    {
                        columns[i % colNum].Add(shapesToArrange[i]); // 按顺序分配到列
                    }

                    // 2. 计算每列的最大宽度
                    List<float> columnWidths = new List<float>();
                    for (int i = 0; i < colNum; i++)
                    {
                        float columnMaxWidth = 0;
                        foreach (var shape in columns[i])
                        {
                            float aspectRatio = shape.Width / shape.Height;

                            if (useCustomWidth && !useCustomHeight)
                            {
                                shape.Width = customWidth;
                                shape.Height = customWidth / aspectRatio;
                            }
                            else if (!useCustomWidth && useCustomHeight)
                            {
                                shape.Height = customHeight;
                                shape.Width = customHeight * aspectRatio;
                            }
                            else if (useCustomWidth && useCustomHeight)
                            {
                                // 取消锁定纵横比 (假设 Shape 类有 LockAspectRatio 属性)
                                // shape.LockAspectRatio = Office.MsoTriState.msoFalse; // 如果使用 Office Interop
                                shape.Width = customWidth;
                                shape.Height = customHeight;
                            }
                            columnMaxWidth = Math.Max(columnMaxWidth, shape.Width);
                        }
                        columnWidths.Add(columnMaxWidth);
                    }
                    float currentX = startX;
                    float rowMaxHeight = 0;
                    int colCount = 0;
                    // 3. 按行进行排列
                    foreach (var shape in shapesToArrange)
                    {
                        float aspectRatio = shape.Width / shape.Height;
                        if (useCustomWidth && !useCustomHeight)
                        {
                            shape.Width = customWidth;
                            shape.Height = customWidth / aspectRatio;
                        }
                        else if (!useCustomWidth && useCustomHeight)
                        {
                            shape.Height = customHeight;
                            shape.Width = customHeight * aspectRatio;
                            // referenceHeight = customHeight;
                            // 需要计算最大占位宽度
                        }
                        else if (useCustomWidth && useCustomHeight)
                        {
                            // 取消锁定纵横比
                            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                            shape.Width = customWidth;
                            shape.Height = customHeight;
                        }

                        if (colCount >= colNum)
                        {
                            colCount = 0;
                            currentX = startX;
                            currentY += rowMaxHeight + rowSpace;
                            rowMaxHeight = 0;
                        }

                        shape.Left = currentX;
                        shape.Top = currentY;
                        rowMaxHeight = Math.Max(rowMaxHeight, shape.Height);
                        currentX += columnWidths[colCount] + colSpace;
                        colCount++;
                    }
                }
                else if (imgAutoAlignAlignTypeDropDown.SelectedItemIndex == 1)
                {
                    // 统一高度排列
                    float referenceHeight = shapesToArrange[0].Height;
                    if (useCustomWidth && !useCustomHeight)
                    {
                        referenceHeight = 0;
                    }
                    float currentX = startX;
                    float rowMaxHeight = 0;
                    int colCount = 0;

                    foreach (var shape in shapesToArrange)
                    {
                        // 保持宽高比调整高度
                        float aspectRatio = shape.Width / shape.Height;
                        if (!useCustomWidth && !useCustomHeight)
                        {
                            shape.Height = referenceHeight;
                            shape.Width = referenceHeight * aspectRatio;
                        }
                        else
                        {
                            if (useCustomWidth && !useCustomHeight)
                            {
                                shape.Width = customWidth;
                                shape.Height = customWidth / aspectRatio;
                            }
                            else if (!useCustomWidth && useCustomHeight)
                            {
                                shape.Height = customHeight;
                                shape.Width = customHeight * aspectRatio;
                                referenceHeight = customHeight;
                            }
                            else
                            {
                                // 取消锁定纵横比
                                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                                shape.Width = customWidth;
                                shape.Height = customHeight;
                            }
                        }

                        if (colCount >= colNum)
                        {
                            colCount = 0;
                            currentX = startX;
                            currentY += referenceHeight + rowSpace;
                            if (useCustomWidth && !useCustomHeight)
                            {
                                referenceHeight = 0;
                            }
                        }

                        shape.Left = currentX;
                        shape.Top = currentY;
                        currentX += shape.Width + colSpace;
                        colCount++;

                        // Calculate the maximum height in the current row
                        if (useCustomWidth && !useCustomHeight)
                        {
                            referenceHeight = Math.Max(referenceHeight, shape.Height);
                        }
                    }
                }
                else
                {
                    // 瀑布流排列：统一所有图片宽度
                    float[] columnTops = new float[colNum];
                    float[] columnLefts = new float[colNum];

                    // 统一所有图片的宽度
                    float uniformWidth = customWidth > 0 ? customWidth : shapesToArrange[0].Width;

                    // 初始化每列的位置
                    for (int i = 0; i < colNum; i++)
                    {
                        columnTops[i] = currentY;
                        columnLefts[i] = startX + i * (uniformWidth + colSpace);
                    }

                    foreach (var shape in shapesToArrange)
                    {
                        // 统一宽度，保持宽高比
                        float aspectRatio = shape.Width / shape.Height;
                        shape.Width = uniformWidth;
                        shape.Height = uniformWidth / aspectRatio;

                        // 找到高度最小的列
                        int minColumn = 0;
                        float minHeight = columnTops[0];
                        for (int i = 1; i < colNum; i++)
                        {
                            if (columnTops[i] < minHeight)
                            {
                                minHeight = columnTops[i];
                                minColumn = i;
                            }
                        }

                        // 放置图片
                        shape.Left = columnLefts[minColumn];
                        shape.Top = columnTops[minColumn];

                        // 更新该列的高度，加上图片高度和行间距
                        columnTops[minColumn] += shape.Height + rowSpace;
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择要对齐的图片。");
            }
        }

        private void gallery1_Click(object sender, RibbonControlEventArgs e) { }

        private void insertCodeBlockButton_Click(object sender, RibbonControlEventArgs e)
        {
            // Create and configure input dialog
            Form inputDialog = new Form()
            {
                Width = 600,
                Height = 400,
                Text = "插入代码块",
                StartPosition = FormStartPosition.CenterScreen, // Center the dialog on the screen
            };

            TextBox codeInput = new TextBox()
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 12),
            };

            ComboBox languageSelect = new ComboBox()
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList,
            };

            // Add common programming languages
            languageSelect.Items.AddRange(
                new string[] { "python", "matlab", "javascript", "html", "css", "R",
                "fortran" }
            );

            // Load default language from settings
            string defaultLanguage = Properties.Settings.Default.selectedCodeLanguage;
            int defaultIndex = languageSelect.Items.IndexOf(defaultLanguage);
            if (defaultIndex >= 0)
            {
                languageSelect.SelectedIndex = defaultIndex;
            }
            else
            {
                languageSelect.SelectedIndex = 0; // Default to first item if not found
            }

            Button okButton = new Button()
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Bottom,
            };

            // Add controls to form
            inputDialog.Controls.AddRange(new Control[] { codeInput, languageSelect, okButton });

            // Show dialog and process result
            if (inputDialog.ShowDialog() == DialogResult.OK)
            {
                string code = codeInput.Text.Trim();
                string language = languageSelect.SelectedItem.ToString();

                // Save selected language to settings
                Properties.Settings.Default.selectedCodeLanguage = language;
                Properties.Settings.Default.Save();

                if (!string.IsNullOrEmpty(code))
                {
                    PowerPoint.Application app = Globals.ThisAddIn.Application;
                    Slide slide = app.ActiveWindow.View.Slide;

                    Shape textBox = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        100,
                        100,
                        500,
                        300
                    );

                    // Set code block style
                    textBox.Fill.Solid();
                    textBox.Fill.ForeColor.RGB = toggleBackgroundCheckBox.Checked
                        ? ColorTranslator.ToOle(Color.FromArgb(30, 30, 30))
                        : ColorTranslator.ToOle(Color.White);
                    textBox.Line.ForeColor.RGB = ColorTranslator.ToOle(
                        Color.FromArgb(200, 200, 200)
                    );
                    textBox.Line.Weight = 1;

                    // Set the code without language markers
                    textBox.TextFrame.TextRange.Text = code;

                    // Apply base formatting
                    textBox.TextFrame.TextRange.Font.Name = "Consolas";
                    textBox.TextFrame.TextRange.Font.Size = 12;
                    textBox.TextFrame.TextRange.Font.Color.RGB = toggleBackgroundCheckBox.Checked
                        ? ColorTranslator.ToOle(Color.White)
                        : ColorTranslator.ToOle(Color.Black);
                    textBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                        PpParagraphAlignment.ppAlignLeft;

                    // Set margins
                    textBox.TextFrame.MarginLeft = 10;
                    textBox.TextFrame.MarginRight = 10;
                    textBox.TextFrame.MarginTop = 5;
                    textBox.TextFrame.MarginBottom = 5;

                    // Apply syntax highlighting
                    var highlighter = new CodeHighlighter(toggleBackgroundCheckBox.Checked);
                    highlighter.ApplyHighlighting(textBox, code, language);

                    // Auto-size the textbox to fit content
                    textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                }
            }
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        // Update background color
                        shape.Fill.Solid();
                        shape.Fill.ForeColor.RGB = toggleBackgroundCheckBox.Checked
                            ? ColorTranslator.ToOle(Color.FromArgb(30, 30, 30))
                            : ColorTranslator.ToOle(Color.White);

                        // Update text color
                        shape.TextFrame.TextRange.Font.Color.RGB = toggleBackgroundCheckBox.Checked
                            ? ColorTranslator.ToOle(Color.White)
                            : ColorTranslator.ToOle(Color.Black);
                    }
                }
            }
        }

        private void insertEquationButton_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide slide = app.ActiveWindow.View.Slide;

            // Prompt user for LaTeX input
            Form inputDialog = new Form()
            {
                Width = 500,
                Height = 500,
                Text = "输入LaTeX公式",
                StartPosition = FormStartPosition.CenterScreen, // Center the dialog on the screen
            };

            TextBox latexInputBox = new TextBox()
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 12),
            };

            Button okButton = new Button()
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Bottom,
            };

            inputDialog.Controls.Add(latexInputBox);
            inputDialog.Controls.Add(okButton);

            if (inputDialog.ShowDialog() == DialogResult.OK)
            {
                string latexInput = latexInputBox.Text.Trim();

                // Remove surrounding $...$, $$...$$, \(...\), \[...\]
                if (latexInput.StartsWith("$") && latexInput.EndsWith("$"))
                {
                    latexInput = latexInput.Trim('$');
                }
                else if (latexInput.StartsWith("$$") && latexInput.EndsWith("$$"))
                {
                    latexInput = latexInput.Trim('$');
                }
                else if (latexInput.StartsWith(@"\(") && latexInput.EndsWith(@"\)"))
                {
                    latexInput = latexInput.Substring(2, latexInput.Length - 4);
                }
                else if (latexInput.StartsWith(@"\[") && latexInput.EndsWith(@"\]"))
                {
                    latexInput = latexInput.Substring(2, latexInput.Length - 4);
                }

                latexInput = latexInput.Replace("\r", "").Replace("\n", ""); // Remove line breaks

                if (!string.IsNullOrEmpty(latexInput))
                {
                    try
                    {
                        // Insert a new textbox in the center of the slide
                        Shape textBox = slide.Shapes.AddTextbox(
                            Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            slide.Master.Width / 2 - 100,
                            slide.Master.Height / 2 - 50,
                            500,
                            500
                        );

                        // Select the newly inserted textbox
                        textBox.Select();
                        app.ActiveWindow.Selection.TextRange.Select();

                        // Run SwitchLatex
                        app.CommandBars.ExecuteMso("EquationInsertNew");
                        Shape equationShape = app.ActiveWindow.Selection.ShapeRange[1];
                        equationShape
                            .TextFrame.TextRange.Characters(
                                1,
                                equationShape.TextFrame.TextRange.Text.Length - 1
                            )
                            .Text = "\u24C9";

                        app.CommandBars.ExecuteMso("EquationInsertNew");
                        app.ActiveWindow.Selection.TextRange.Select();
                        Shape equationShape2 = app.ActiveWindow.Selection.ShapeRange[1];
                        // Set the LaTeX input to the equation shape
                        equationShape2
                            .TextFrame.TextRange.Characters(
                                1,
                                equationShape2.TextFrame.TextRange.Text.Length - 1
                            )
                            .Text = latexInput;

                        // Convert to professional format
                        app.CommandBars.ExecuteMso("EquationProfessional");

                        textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("An error occurred: " + ex.Message);
                    }
                }
            }
        }

        private int GetActualPosition(string text, int position)
        {
            return position - text.Substring(0, position).Count(c => c == '\r');
        }

        private string ConvertMarkdownToHtml(string markdown)
        {
            try
            {
                var html = Markdown.ToHtml(markdown);
                //MessageBox.Show($"Markdown转换: {html}");
                return html;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Markdown转换错误: {ex.Message}");
                return markdown; // 转换失败时返回原文本
            }
        }
        private void insertMarkdown_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Form inputDialog = new Form
                {
                    Width = 600,
                    Height = 400,
                    Text = "插入Markdown",
                    StartPosition = FormStartPosition.CenterScreen,
                };

                TextBox markdownInput = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Dock = DockStyle.Fill,
                    Font = new Font("Consolas", 12),
                };

                Button okButton = new Button
                {
                    Text = "确定",
                    DialogResult = DialogResult.OK,
                    Dock = DockStyle.Bottom,
                };

                inputDialog.Controls.Add(markdownInput);
                inputDialog.Controls.Add(okButton);

                DialogResult result = inputDialog.ShowDialog();

                if (result == DialogResult.OK)
                {
                    string markdown = markdownInput.Text?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(markdown))
                    {
                        Slide slide = app.ActiveWindow.View.Slide;

                        // Split markdown into segments
                        var segments = SplitMarkdownIntoSegments(markdown);

                        float currentTop = slide.Master.Height / 2; // Starting position
                        float left = (slide.Master.Width - 500) / 2; // Center horizontally

                        foreach (var segment in segments)
                        {
                            try
                            {
                                Shape shape = null;
                                if (segment.IsCodeBlock)
                                {
                                    shape = InsertCodeBlock(
                                        segment.Content,
                                        segment.Language,
                                        left,
                                        currentTop
                                    );
                                }
                                else if (segment.IsTable)
                                {
                                    shape = InsertTable(segment.Content, left, currentTop);
                                }
                                else if (segment.IsMathBlock)
                                {
                                    shape = InsertMathBlock(segment.Content, left, currentTop);
                                }
                                else if (segment.IsBlockQuote)
                                {
                                    shape = InsertBlockQuote(segment.Content, left, currentTop);
                                }
                                else
                                {
                                    string html = ProcessMarkdown(segment.Content);
                                    if (!string.IsNullOrEmpty(html))
                                    {
                                        // Add retry mechanism for clipboard operations
                                        int retryCount = 3;
                                        while (retryCount > 0)
                                        {
                                            try
                                            {
                                                CopyHtmlToClipBoard(segment.Content, html);
                                                System.Threading.Thread.Sleep(100); // Add 100ms delay
                                                ShapeRange textContent = slide.Shapes.Paste();

                                                if (textContent != null && textContent.Count > 0)
                                                {
                                                    Shape textShape = textContent[1];
                                                    textShape.Width = 500;
                                                    textShape.Left = left;
                                                    textShape.Top = currentTop;
                                                    currentTop += textShape.Height + 10;

                                                    // Process inline math formulas
                                                    ProcessInlineMathFormulas(textShape);

                                                    if (
                                                        textShape.TextFrame.HasText
                                                        == Office.MsoTriState.msoTrue
                                                    )
                                                    {
                                                        TextRange textRange = textShape
                                                            .TextFrame
                                                            .TextRange;
                                                        foreach (
                                                            TextRange paragraph in textRange.Paragraphs(
                                                                -1
                                                            )
                                                        ) // Changed this line
                                                        {
                                                            if (
                                                                paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .Type
                                                                != PpBulletType.ppBulletNone
                                                            )
                                                            {
                                                                // 保存列表样式
                                                                PpBulletType ppBulletType =
                                                                    paragraph
                                                                        .ParagraphFormat
                                                                        .Bullet
                                                                        .Type;
                                                                int character = paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .Character;
                                                                int startValue = paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .StartValue; // 有序列表的编号
                                                                PpNumberedBulletStyle stype =
                                                                    paragraph
                                                                        .ParagraphFormat
                                                                        .Bullet
                                                                        .Style; // 有序列表的样式：1、A、一等

                                                                // 重新设置列表样式，曲线救国来添加悬挂缩进（找不到代码的方式直接添加悬挂缩进
                                                                //paragraph.ParagraphFormat.Bullet.Type = PpBulletType.ppBulletNone;
                                                                paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .Type = ppBulletType;
                                                                paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .Character = character;
                                                                if (
                                                                    ppBulletType
                                                                    == PpBulletType.ppBulletNumbered
                                                                )
                                                                {
                                                                    paragraph
                                                                        .ParagraphFormat
                                                                        .Bullet
                                                                        .StartValue = startValue;
                                                                    paragraph
                                                                        .ParagraphFormat
                                                                        .Bullet
                                                                        .Style = stype;
                                                                }
                                                                // 列表样式不受后面字体样式的干扰
                                                                paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .UseTextFont = Office
                                                                    .MsoTriState
                                                                    .msoFalse;
                                                                paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .UseTextColor = Office
                                                                    .MsoTriState
                                                                    .msoFalse;
                                                                paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .Font
                                                                    .Bold = Office
                                                                    .MsoTriState
                                                                    .msoFalse;
                                                                paragraph
                                                                    .ParagraphFormat
                                                                    .Bullet
                                                                    .Font
                                                                    .Italic = Office
                                                                    .MsoTriState
                                                                    .msoFalse;
                                                                // 弹窗输出ppBulletType是什么
                                                                // MessageBox.Show($"Bullet type: {ppBulletType}, Start value: {startValue}, Character: {character}, Style: {stype}");

                                                                string text = paragraph.Text.Trim();
                                                                if (text.StartsWith("- [x]"))
                                                                {
                                                                    char myCharacter = (char)9745; // ☑
                                                                    paragraph
                                                                        .ParagraphFormat
                                                                        .Bullet
                                                                        .Character = myCharacter;
                                                                    paragraph.Text = text.Substring(
                                                                            5
                                                                        )
                                                                        .Trim(); // Remove "- [x]"
                                                                }
                                                                else if (text.StartsWith("- [ ]"))
                                                                {
                                                                    char myCharacter = (char)9744; // ☐
                                                                    paragraph
                                                                        .ParagraphFormat
                                                                        .Bullet
                                                                        .Character = myCharacter;
                                                                    paragraph.Text = text.Substring(
                                                                            5
                                                                        )
                                                                        .Trim(); // Remove "- [ ]"
                                                                }
                                                            }
                                                        }
                                                    }
                                                    break; // Success, exit retry loop
                                                }
                                            }
                                            catch (System.Runtime.InteropServices.COMException)
                                            {
                                                retryCount--;
                                                if (retryCount <= 0)
                                                {
                                                    MessageBox.Show(
                                                        $"无法粘贴内容: {segment.Content.Substring(0, Math.Min(30, segment.Content.Length))}..."
                                                    );
                                                }
                                                System.Threading.Thread.Sleep(200); // Wait longer before retry
                                            }
                                        }
                                    }
                                }

                                if (shape != null)
                                {
                                    currentTop += shape.Height + 10;
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"处理段落时出错: {ex.Message}");
                                continue; // Continue with next segment
                            }
                        }
                        inputDialog.Dispose();
                        Clipboard.Clear();
                        var dataObject = new DataObject();
                        dataObject.SetData(DataFormats.UnicodeText, markdown);
                        Clipboard.SetDataObject(dataObject, true, 3, 100); // Add retry and timeout parameters
                    }
                }

                inputDialog.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作过程中出错: {ex.Message}\n\n{ex.StackTrace}");
            }
        }


        private void ProcessInlineMathFormulas(Shape textShape)
        {
            TextRange textRange = textShape.TextFrame.TextRange;
            string text = textRange.Text;
            // Regex pattern to find math expressions between $ signs
            var matches = Regex.Matches(text, @"\$([^$\n]+?)\$");
            // matches.Count如果=0，说明没有匹配到，直接返回
            if (matches.Count == 0)
            {
                return;
            }
            // 创建tempShape，如果不创建，行内数学公式包括分式就不会正常转化
            Shape tempShape = InsertMathBlock("a", 0, 0);
            // 删除mathShape
            tempShape.Delete();

            // Process matches in reverse order to maintain correct indices
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                var match = matches[i];
                int start = match.Index + 1; // 1-based start index of the match (e.g., the first '$')
                int length = match.Length; // Length of the matched string (e.g., "$formula$")
                string formula = match.Groups[1].Value; // Content within $...$ (e.g., "formula")

                // Select the range "$formula$"
                TextRange selectedRange = textRange.Characters(start, length);
                // Replace its text with "formula"
                selectedRange.Text = formula;
                selectedRange.Select();
                app.CommandBars.ExecuteMso("EquationInsertNew");

                app.CommandBars.ExecuteMso("EquationProfessional");
            }
        }

        private Shape InsertCodeBlock(string code, string language, float left, float top)
        {
            Slide slide = app.ActiveWindow.View.Slide;
            Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                500,
                300
            );

            // Set code block style
            textBox.Fill.Solid();
            textBox.Fill.ForeColor.RGB = toggleBackgroundCheckBox.Checked
                ? ColorTranslator.ToOle(Color.FromArgb(30, 30, 30))
                : ColorTranslator.ToOle(Color.White);
            textBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
            textBox.Line.Weight = 1;

            textBox.TextFrame.TextRange.Text = code;

            // Apply base formatting
            textBox.TextFrame.TextRange.Font.Name = "Consolas";
            textBox.TextFrame.TextRange.Font.Size = 12;
            textBox.TextFrame.TextRange.Font.Color.RGB = toggleBackgroundCheckBox.Checked
                ? ColorTranslator.ToOle(Color.White)
                : ColorTranslator.ToOle(Color.Black);
            textBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                PpParagraphAlignment.ppAlignLeft;

            // Set margins
            textBox.TextFrame.MarginLeft = 10;
            textBox.TextFrame.MarginRight = 10;
            textBox.TextFrame.MarginTop = 5;
            textBox.TextFrame.MarginBottom = 5;

            // Apply syntax highlighting
            var highlighter = new CodeHighlighter(toggleBackgroundCheckBox.Checked);
            highlighter.ApplyHighlighting(textBox, code, language);

            // Auto-size the textbox to fit content
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

            return textBox;
        }

        public class ImageGroup
        {
            public List<Shape> Shapes { get; set; } = new List<Shape>();
            public float MinTop { get; set; }
            public float MaxBottom { get; set; }

            public bool OverlapsWith(Shape shape)
            {
                float shapeHeight = shape.Height;
                float threshold = shapeHeight * 0.5f; // 50% of shape height
                float shapeBottom = shape.Top + shapeHeight;

                // Calculate overlap height
                float overlapStart = Math.Max(MinTop, shape.Top);
                float overlapEnd = Math.Min(MaxBottom, shapeBottom);
                float overlapHeight = overlapEnd - overlapStart;

                return overlapHeight >= threshold;
            }

            public void AddShape(Shape shape)
            {
                if (Shapes.Count == 0)
                {
                    MinTop = shape.Top;
                    MaxBottom = shape.Top + shape.Height;
                }
                else
                {
                    MinTop = Math.Min(MinTop, shape.Top);
                    MaxBottom = Math.Max(MaxBottom, shape.Top + shape.Height);
                }
                Shapes.Add(shape);
            }
        }

        private class MarkdownSegment
        {
            public string Content { get; set; }
            public bool IsCodeBlock { get; set; }
            public bool IsTable { get; set; }
            public bool IsMathBlock { get; set; }
            public bool IsBlockQuote { get; set; } // Add this line
            public string Language { get; set; }
        }

        private List<MarkdownSegment> SplitMarkdownIntoSegments(string markdown)
        {
            var segments = new List<MarkdownSegment>();
            var currentPosition = 0;

            // Updated pattern to better handle tables
            // 1. Tables must start with a header line
            // 2. Followed by a separator line
            // 3. Then one or more data lines
            var pattern =
                @"(?:```(\w*)\r?\n(.*?)\r?\n```)|"
                + // Code blocks
                @"(?:\|[^\n]*\|\r?\n\|[-|\s]*\|\r?\n(?:\|[^\n]*\|\r?\n)*\|[^\n]*\|?)|"
                + // Tables
                @"(\$\$[\s\S]*?\$\$)|"
                + // Math blocks
                @"(?:(?:^|\n)(?:>[^\n]*(?:\r?\n>[^\n]*)*))"; // 引述块（修改后的模式）

            var regex = new Regex(pattern, RegexOptions.Multiline | RegexOptions.Singleline);

            var matches = regex.Matches(markdown);

            foreach (Match match in matches)
            {
                // Add text before special block if exists
                if (match.Index > currentPosition)
                {
                    string textBefore = markdown.Substring(
                        currentPosition,
                        match.Index - currentPosition
                    );
                    if (!string.IsNullOrWhiteSpace(textBefore))
                    {
                        segments.Add(
                            new MarkdownSegment
                            {
                                Content = textBefore.Trim(),
                                IsCodeBlock = false,
                                IsTable = false,
                                IsMathBlock = false,
                                IsBlockQuote = false,
                            }
                        );
                    }
                }

                string content = match.Value;

                // Determine block type and add segment
                if (content.StartsWith("```"))
                {
                    var lines = content.Split(
                        new[] { '\r', '\n' },
                        StringSplitOptions.RemoveEmptyEntries
                    );
                    var language = lines[0].Substring(3).Trim();
                    var codeContent = string.Join("\n", lines.Skip(1).Take(lines.Length - 2));

                    segments.Add(
                        new MarkdownSegment
                        {
                            Content = codeContent,
                            Language = string.IsNullOrEmpty(language) ? "text" : language,
                            IsCodeBlock = true,
                            IsTable = false,
                            IsMathBlock = false,
                            IsBlockQuote = false,
                        }
                    );
                }
                else if (content.StartsWith("|"))
                {
                    // Clean up table content (remove trailing whitespace and newlines)
                    content = string.Join(
                        "\n",
                        content
                            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(line => line.Trim())
                            .Where(line => line.StartsWith("|") && line.EndsWith("|"))
                    );

                    segments.Add(
                        new MarkdownSegment
                        {
                            Content = content,
                            IsCodeBlock = false,
                            IsTable = true,
                            IsMathBlock = false,
                            IsBlockQuote = false,
                        }
                    );
                }
                else if (content.StartsWith("$$"))
                {
                    content = string.Join(
                        "\n",
                        content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    );
                    
                    // Remove surrounding $$ markers like in insertEquationButton_Click
                    string mathContent = content.Replace("\n", ""); // Remove line breaks
                    if (mathContent.StartsWith("$$") && mathContent.EndsWith("$$"))
                    {
                        mathContent = mathContent.Substring(2, mathContent.Length - 4);
                    }
                    
                    segments.Add(
                        new MarkdownSegment
                        {
                            Content = mathContent,
                            IsCodeBlock = false,
                            IsTable = false,
                            IsMathBlock = true,
                            IsBlockQuote = false,
                        }
                    );
                }
                else if (content.TrimStart('\r', '\n').StartsWith(">"))
                {
                    // Clean up block quote content
                    content = string.Join(
                        "\n",
                        content
                            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(line => line.TrimStart('>', ' '))
                    );

                    segments.Add(
                        new MarkdownSegment
                        {
                            Content = content,
                            IsCodeBlock = false,
                            IsTable = false,
                            IsMathBlock = false,
                            IsBlockQuote = true,
                        }
                    );
                }

                currentPosition = match.Index + match.Length;
            }

            // Add remaining text if exists
            if (currentPosition < markdown.Length)
            {
                string remainingText = markdown.Substring(currentPosition);
                if (!string.IsNullOrWhiteSpace(remainingText))
                {
                    segments.Add(
                        new MarkdownSegment
                        {
                            Content = remainingText.Trim(),
                            IsCodeBlock = false,
                            IsTable = false,
                            IsMathBlock = false,
                            IsBlockQuote = false,
                        }
                    );
                }
            }

            return segments;
        }

        private Shape InsertTable(string tableContent, float left, float top)
        {
            Slide slide = app.ActiveWindow.View.Slide;
            Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                500,
                300
            );

            // Convert markdown table to HTML
            // Configure the pipeline with all advanced extensions active
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            string html = Markdown.ToHtml(tableContent, pipeline);
            html = html.Replace(
                "<table>",
                "<table style='width:500px; border-collapse:collapse;border:1pt solid black;'>"
            );
            html = html.Replace("<td>", "<td style='border:1pt solid black;'>");
            html = html.Replace("<th>", "<th style='border:1pt solid black;'>");

            // Create a temporary DataObject for the table content

            CopyHtmlToClipBoard(tableContent, html);
            System.Threading.Thread.Sleep(100);

            ShapeRange tableShape = slide.Shapes.Paste();
            if (tableShape != null && tableShape.Count > 0)
            {
                tableShape[1].Left = left;
                tableShape[1].Top = top;
                textBox.Delete();
                return tableShape[1];
            }

            return textBox;
        }

        private Shape InsertMathBlock(string mathContent, float left, float top)
        {
            Slide slide = app.ActiveWindow.View.Slide;

            // Insert a new textbox
            Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                500,
                500
            );

            // Select the newly inserted textbox
            textBox.Select();
            app.ActiveWindow.Selection.TextRange.Select();

            // Run SwitchLatex
            app.CommandBars.ExecuteMso("EquationInsertNew");
            Shape equationShape = app.ActiveWindow.Selection.ShapeRange[1];
            equationShape
                .TextFrame.TextRange.Characters(
                    1,
                    equationShape.TextFrame.TextRange.Text.Length - 1
                )
                .Text = "\u24C9";

            app.CommandBars.ExecuteMso("EquationInsertNew");
            app.ActiveWindow.Selection.TextRange.Select();
            Shape equationShape2 = app.ActiveWindow.Selection.ShapeRange[1];
            // Set the LaTeX input to the equation shape
            equationShape2
                .TextFrame.TextRange.Characters(
                    1,
                    equationShape2.TextFrame.TextRange.Text.Length - 1
                )
                .Text = mathContent;

            // Convert to professional format
            app.CommandBars.ExecuteMso("EquationProfessional");
            // Auto-size and position
            equationShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            equationShape.Left = left;
            equationShape.Top = top;

            return equationShape;
        }

        private Shape InsertBlockQuote(string content, float left, float top)
        {
            Slide slide = app.ActiveWindow.View.Slide;
            Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                500,
                300
            );

            // Configure Markdown pipeline
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

            // Convert to HTML and remove blockquote tags
            string html = Markdown
                .ToHtml(content, pipeline)
                .Replace("<blockquote>", "")
                .Replace("</blockquote>", "");

            // Add custom styling
            html = $"<div style='font-family: 微软雅黑; padding: 10px;'>{html}</div>";

            // Copy to clipboard and paste
            CopyHtmlToClipBoard(content, html);
            System.Threading.Thread.Sleep(100);

            ShapeRange quoteShape = slide.Shapes.Paste();
            if (quoteShape != null && quoteShape.Count > 0)
            {
                Shape shape = quoteShape[1];
                shape.Left = left;
                shape.Top = top;

                // Add black border
                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                shape.Line.Weight = 1;

                textBox.Delete();
                return shape;
            }

            return textBox;
        }

        private string ProcessMarkdown(string markdown)
        {
            var codeBlockRegex = new Regex(@"```.*?\r?\n(.*?)\r?\n```", RegexOptions.Singleline);

            markdown = codeBlockRegex.Replace(markdown, string.Empty);

            // Convert remaining markdown to HTML
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            string html = Markdown.ToHtml(markdown, pipeline);

            // Add checkbox markers after the checkboxes
            html = html.Replace(
                "<input disabled=\"disabled\" type=\"checkbox\" checked=\"checked\" />",
                "- [x]"
            );
            html = html.Replace("<input disabled=\"disabled\" type=\"checkbox\" />", "- [ ]");
            // Add table styling
            html = html.Replace(
                "<table>",
                "<table style='width:500px; border-collapse:collapse;border:1pt solid黑色;'>"
            );
            html = html.Replace("<td>", "<td style='border:1pt solid black;'>");
            html = html.Replace("<th>", "<th style='border:1pt solid black;'>");

            html = html.Replace("<li>", "<li style='margin-left: 10px;'>");
            html = html.Replace("<code>", "<span style='color: #C00000; font-family: Consolas;'>");
            html = html.Replace("</code>", "</span>");

            html = $"<div style='font-family: 微软雅黑;'>{html}</div>";

            // 把<span class="math">\(...\)</span>转换$...$
            html = Regex.Replace(
                html,
                @"<span class=""math"">\\\((.+?)\\\)</span>",
                m => $"${m.Groups[1].Value}$"
            );
            // 弹窗显示
            //MessageBox.Show($"Markdown转换: {html}");
            return html;
        }

        public void CopyHtmlToClipBoard(string markdown, string html)
        {
            try
            {
                var utf = Encoding.UTF8;
                var format =
                    "Version:0.9\r\nStartHTML:{0:000000}\r\nEndHTML:{1:000000}\r\nStartFragment:{2:000000}\r\nEndFragment:{3:000000}\r\n";
                var text =
                    "<html>\r\n<head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset="
                    + utf.WebName
                    + "\">\r\n<title>HTML clipboard</title>\r\n</head>\r\n<body>\r\n<!--StartFragment-->";
                var text2 = "<!--EndFragment-->\r\n</body>\r\n</html>\r\n";
                var s = string.Format(format, 0, 0, 0, 0);
                var byteCount = utf.GetByteCount(s);
                var byteCount2 = utf.GetByteCount(text);
                var byteCount3 = utf.GetByteCount(html);
                var byteCount4 = utf.GetByteCount(text2);
                var s2 =
                    string.Format(
                        format,
                        byteCount,
                        byteCount + byteCount2 + byteCount3 + byteCount4,
                        byteCount + byteCount2,
                        byteCount + byteCount2 + byteCount3
                    )
                    + text
                    + html
                    + text2;

                var dataObject = new DataObject();
                dataObject.SetData(DataFormats.Html, new MemoryStream(utf.GetBytes(s2)));
                dataObject.SetData(DataFormats.UnicodeText, markdown);

                int retryCount = 3;
                while (retryCount > 0)
                {
                    try
                    {
                        Clipboard.SetDataObject(dataObject, true, 3, 100); // Add retry and timeout parameters
                        break;
                    }
                    catch (Exception)
                    {
                        retryCount--;
                        if (retryCount <= 0)
                            throw;
                        System.Threading.Thread.Sleep(100);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制到剪贴板时出错: {ex.Message}");
                throw;
            }
        }

        private void copyCrop_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape shape = sel.ShapeRange[1];

                // 保存裁剪设置
                cropLeft = shape.PictureFormat.CropLeft;
                cropRight = shape.PictureFormat.CropRight;
                cropTop = shape.PictureFormat.CropTop;
                cropBottom = shape.PictureFormat.CropBottom;

                // 保存原始高度
                currentCropedHeight = shape.Height;
                float croppedPixels = cropTop + cropBottom;
                originalHeight = currentCropedHeight + croppedPixels;

                hasCopiedCrop = true;
                //MessageBox.Show("已复制图片裁剪设置");
            }
            else
            {
                MessageBox.Show("请选择一个图片对象");
            }
        }

        private void pasteCrop_Click(object sender, RibbonControlEventArgs e)
        {
            if (!hasCopiedCrop)
            {
                MessageBox.Show("请先复制图片裁剪设置");
                return;
            }
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
                {
                    try
                    {
                        // Store original position
                        float originalLeft = shape.Left;
                        float originalTop = shape.Top;

                        // Clear existing crop settings
                        shape.PictureFormat.CropLeft = 0;
                        shape.PictureFormat.CropRight = 0;
                        shape.PictureFormat.CropTop = 0;
                        shape.PictureFormat.CropBottom = 0;

                        // Restore to original height
                        shape.Height = originalHeight;

                        // Apply crop settings
                        shape.PictureFormat.CropLeft = cropLeft;
                        shape.PictureFormat.CropRight = cropRight;
                        shape.PictureFormat.CropTop = cropTop;
                        shape.PictureFormat.CropBottom = cropBottom;

                        shape.Height = currentCropedHeight;

                        // Restore original position
                        shape.Left = originalLeft;
                        shape.Top = originalTop;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"应用裁剪设置时出错: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择要应用裁剪设置的图片");
            }
        }

        private void openGithub_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/Achuan-2/my_ppt_plugin/");
        }

        private void openDoc_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start(
                "https://www.yuque.com/achuan-2/blog/etzcergpmb4rr2sk/"
            );
        }

        private void current_Version(object sender, RibbonControlEventArgs e)
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            Version version = assembly.GetName().Version;
            MessageBox.Show($"Version {version}", "Current Version");
        }

        private void aboutDeveloper_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                "开发者: Achuan-2\n邮箱: achuan-2@outlook.com\nGithub地址：https://github.com/Achuan-2",
                "关于开发者"
            );
        }

        private void positionSortCheckBox_Click(object sender, RibbonControlEventArgs e) { }

        private void addLabelsButton_Click(object sender, RibbonControlEventArgs e)
        {
            string fontFamily = labelFontNameEditBox.Text; // 修改为使用新控件
            float fontSize;
            if (!float.TryParse(labelFontSizeEditBox.Text, out fontSize)) // 修改为使用新控件
            {
                MessageBox.Show("请输入有效的字体大小。");
                return;
            }
            float labelOffsetX;
            if (!float.TryParse(labelOffsetXEditBox.Text, out labelOffsetX))
            {
                MessageBox.Show("请输入有效的X偏移量。");
                return;
            }
            float labelOffsetY;
            if (!float.TryParse(labelOffsetYEditBox.Text, out labelOffsetY))
            {
                MessageBox.Show("请输入有效的Y偏移量。");
                return;
            }
            string labelTemplate = labelTemplateComboBox.Text;

            // 获取起始编号
            int startIndex = 1;
            if (!string.IsNullOrEmpty(labelIndex.Text) && !int.TryParse(labelIndex.Text, out startIndex))
            {
                MessageBox.Show("请输入有效的起始编号。");
                return;
            }

            AddLabelsToImages(fontFamily, fontSize, labelOffsetX, labelOffsetY, labelTemplate, startIndex, true);
        }

        /// <summary>
        /// 图片添加标签
        /// </summary>
        /// <param name="fontFamily"></param>
        /// <param name="fontSize"></param>
        /// <param name="labelOffsetX"></param>
        /// <param name="labelOffsetY"></param>
        /// <param name="labelTemplate">标签格式</param>
        /// <param name="startIndex">起始编号</param>
        /// <param name="isAddLabels">是否为添加标签模式，true为添加，false为更新</param>
        private void AddLabelsToImages(
            string fontFamily,
            float fontSize,
            float labelOffsetX,
            float labelOffsetY,
            string labelTemplate,
            int startIndex = 1,
            bool isAddLabels = true
        )
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionShapes)
            {
                // 需要弹窗显示sel.Type
                MessageBox.Show($"请选择要添加标签的图片。当前选择类型: {sel.Type}");
                // MessageBox.Show("请选择要添加标签的图片。");
                return;
            }

            var templates = new Dictionary<string, string>
            {
                { "A", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" },
                { "a", "abcdefghijklmnopqrstuvwxyz" },
                { "A)", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" },
                { "a)", "abcdefghijklmnopqrstuvwxyz" },
                { "(A)", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" },
                { "(a)", "abcdefghijklmnopqrstuvwxyz" },
                { "1", "123456789" }, // Added numeric template
                { "1)", "123456789" }, // Added numeric template with parenthesis
                { "Ⅰ", "ⅠⅡⅢⅣⅤⅥⅦⅦⅨⅩ" },
                { "Ⅰ)", "ⅠⅡⅢⅣⅤⅥⅦⅦⅨⅩ" },
                { "①", "①②③④⑤⑥⑦⑧⑨⑩" },
                { "①)", "①②③④⑤⑥⑦⑧⑨⑩" },
                { "一", "一二三四五六七八九十" },
                { "一)", "一二三四五六七八九十" },
            };

            if (!templates.ContainsKey(labelTemplate))
            {
                labelTemplate = "A";
            }

            string labels = templates[labelTemplate];
            bool isNumeric = labelTemplate.StartsWith("1");
            int selectionCount = sel.ShapeRange.Count;

            // Create groups based on vertical position
            var groups = new List<ImageGroup>();
            var selectedImgShapes = new List<Shape>();
            foreach (Shape shape in sel.ShapeRange)
            {
                // // Skip text boxes if excludeTextcheckBox is checked
                // if (
                //     shape.Type == Office.MsoShapeType.msoTextBox
                //     || shape.Type == Office.MsoShapeType.msoAutoShape
                // )
                // {
                //     continue;
                // }
                selectedImgShapes.Add(shape);
            }
            if (selectedImgShapes.Count == 0)
            {
                MessageBox.Show("请选择要添加标签的图片。");
                return;
            }

            // Group shapes based on vertical overlap
            foreach (var shape in selectedImgShapes)
            {
                bool addedToExistingGroup = false;
                foreach (var group in groups)
                {
                    if (group.OverlapsWith(shape))
                    {
                        group.AddShape(shape);
                        addedToExistingGroup = true;
                        break;
                    }
                }

                if (!addedToExistingGroup)
                {
                    var newGroup = new ImageGroup();
                    newGroup.AddShape(shape);
                    groups.Add(newGroup);
                }
            }

            // Sort shapes within each group by x position
            foreach (var group in groups)
            {
                group.Shapes.Sort((a, b) => a.Left.CompareTo(b.Left));
            }

            // Sort groups by MinTop
            groups.Sort((a, b) => a.MinTop.CompareTo(b.MinTop));

            // Create flattened list of sorted shapes
            var sortedShapes = new List<Shape>();
            foreach (var group in groups)
            {
                sortedShapes.AddRange(group.Shapes);
            }

            // Add labels to sorted shapes
            for (int i = 0; i < sortedShapes.Count; i++)
            {
                try
                {
                    var item = sortedShapes[i];
                    string label;
                    if (isNumeric)
                    {
                        label = (startIndex + i).ToString();
                    }
                    else
                    {
                        int labelIndex = (startIndex - 1 + i) % labels.Length;
                        label = labels[labelIndex].ToString();
                    }

                    if (labelTemplate.EndsWith(")") && !labelTemplate.StartsWith("("))
                    {
                        label += ")";
                    }
                    else if (labelTemplate.StartsWith("(") && labelTemplate.EndsWith(")"))
                    {
                        label = "(" + label + ")";
                    }

                    var textBox = app.ActiveWindow.View.Slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        item.Left + labelOffsetX,
                        item.Top + labelOffsetY,
                        0, // Initial width
                        fontSize * 2
                    );

                    // Set the text and font properties
                    textBox.TextFrame.TextRange.Text = label;
                    textBox.TextFrame.TextRange.Font.Size = fontSize;
                    textBox.TextFrame.TextRange.Font.Name = fontFamily;
                    textBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                        PpParagraphAlignment.ppAlignLeft;

                    // Auto-size the textbox to fit the text
                    textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                    // 不自动换行
                    textBox.TextFrame.WordWrap = Office.MsoTriState.msoFalse;

                    // 自动加粗
                    if (labelBoldcheckBox.Checked)
                    {
                        textBox.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    }
                    // 自动选择
                    if (i == 0)
                    {
                        textBox.Select(Office.MsoTriState.msoTrue);
                    }
                    else
                    {
                        textBox.Select(Office.MsoTriState.msoFalse);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"添加标签时出错: {ex.Message}");
                }
            }
            
            // 如果开启了编号自动更新，并且是添加标签模式，则更新起始编号
            if (isAddLabels && labelIndexUpdatecheckBox.Checked)
            {
                int nextIndex = startIndex + sortedShapes.Count;
                labelIndex.Text = nextIndex.ToString();
            }
        }

        private static PowerPoint.Font _copiedFont;

        private void copyStyle_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape sourceShape = sel.ShapeRange[1];
                // 捕获格式
                sourceShape.PickUp();
            }
            else if (sel.Type == PpSelectionType.ppSelectionText)
            {
                _copiedFont = sel.TextRange.Font;
            }
            else { }
        }

        private void pasteStyle_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
                {
                    shape.Apply();
                }
            }
            else if (sel.Type == PpSelectionType.ppSelectionText)
            {
                ApplyFont(sel.TextRange.Font);
            }
            else { }
        }

        private void ApplyFont(PowerPoint.Font targetFont)
        {
            targetFont.Name = _copiedFont.Name;
            targetFont.Size = _copiedFont.Size;
            targetFont.Bold = _copiedFont.Bold;
            targetFont.Italic = _copiedFont.Italic;
            targetFont.Color.RGB = _copiedFont.Color.RGB;
            targetFont.Underline = _copiedFont.Underline;
        }

        private void pastePictureAndText(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                try
                {
                    // Store original position
                    // float left = sel.ShapeRange.Left;
                    // float top = sel.ShapeRange.Top;

                    // Group the shapes first if multiple shapes selected
                    Shape groupedShape;
                    try
                    {
                        // First attempt - try to group directly
                        groupedShape =
                            sel.ShapeRange.Count > 1 ? sel.ShapeRange.Group() : sel.ShapeRange[1];
                    }
                    catch (Exception ex)
                    {
                        // If direct grouping fails, try the copy-delete-paste-group approach
                        try
                        {
                            // Copy the shapes
                            sel.ShapeRange.Copy();

                            // Delete original shapes
                            sel.ShapeRange.Delete();

                            // Paste back the shapes
                            ShapeRange pastedShapes2 = app.ActiveWindow.View.Slide.Shapes.Paste();

                            // Try grouping again
                            groupedShape =
                                pastedShapes2.Count > 1 ? pastedShapes2.Group() : pastedShapes2[1];
                        }
                        catch (Exception innerEx)
                        {
                            MessageBox.Show(
                                $"无法组合对象: {innerEx.Message}\n原始错误: {ex.Message}"
                            );
                            return;
                        }
                    }

                    // Copy grouped shape
                    groupedShape.Copy();

                    // Delete original shape
                    groupedShape.Delete();

                    // Paste as Enhanced Metafile
                    ShapeRange pastedShapes = app.ActiveWindow.View.Slide.Shapes.PasteSpecial(
                        PpPasteDataType.ppPasteEnhancedMetafile
                    );

                    // // Move to original position
                    // if (pastedShapes != null)
                    // {
                    //     pastedShapes.Left = left;
                    //     pastedShapes.Top = top;
                    //     pastedShapes.Select();
                    // }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"粘贴为增强型图形时出错: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("请先选择要转换的对象。");
            }
        }

        private void imgAutoAlign_rowSpace_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string str1 = imgAutoAlign_rowSpace.Text.Split(new char[] { '≈' })[1];
            if (str1 != null)
            {
                fontSizeEditBox.Text = Regex.Replace(str1, @"[^\d.\d]", "");
            }
            AlignPics();
        }

        private void imgAutoAlign_colNum_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        private void imgAutoAlign_colSpace_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        private void imgWidthEditBpx_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        private void imgHeightEditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        private void excludeTextcheckBox_Click(object sender, RibbonControlEventArgs e) { }

        private void excludeTextcheckBox2_Click(object sender, RibbonControlEventArgs e) { }

        private void donate(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.yuque.com/achuan-2");
        }

        private void developer_website(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.github.com/achuan-2");
        }
        public void ExportOriginalImage_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 1. 获取当前PowerPoint应用实例和选中的对象
                var app = Globals.ThisAddIn.Application;
                var activeWindow = app.ActiveWindow;

                if (app.ActivePresentation == null)
                {
                    MessageBox.Show("请先打开一个演示文稿。", "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 检查是否选中了图形
                if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("请先选择一个图片。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 检查是否只选中了一个图形
                var shapeRange = activeWindow.Selection.ShapeRange;
                if (shapeRange.Count != 1)
                {
                    MessageBox.Show("请只选择一个图片进行导出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var shape = shapeRange[1];

                // 检查选中的是否是图片类型


                // 2. 获取必要信息：Shape ID, Slide Object 和演示文稿路径
                uint shapeId = (uint)shape.Id;
                Slide vstoSlide = shape.Parent;
                uint slideIdValue = (uint)vstoSlide.SlideID; // 获取幻灯片的唯一ID

                // 保存演示文稿以确保图片文件嵌入正确
                app.ActivePresentation.Save();
                string presentationPath = app.ActivePresentation.FullName;

                // 确保演示文稿已保存
                if (string.IsNullOrEmpty(presentationPath) || !File.Exists(presentationPath))
                {
                    MessageBox.Show("请先保存当前演示文稿再执行导出操作。", "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 3. 使用 Open XML SDK 进行操作
                using (PresentationDocument presDoc = PresentationDocument.Open(presentationPath, false)) // false = read-only
                {
                    PresentationPart presPart = presDoc.PresentationPart;
                    if (presPart == null)
                    {
                        MessageBox.Show("无法加载演示文稿的核心部分。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 4. 通过 SlideID 查找对应的 SlideId 条目，从而获取其 RelationshipId
                    var slideIdEntry = presPart.Presentation.SlideIdList.ChildElements
                        .OfType<DocumentFormat.OpenXml.Presentation.SlideId>()
                        .FirstOrDefault(s => s.Id != null && s.Id.Value == slideIdValue);

                    if (slideIdEntry == null || slideIdEntry.RelationshipId == null)
                    {
                        MessageBox.Show("无法在文档结构中定位到当前幻灯片。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 使用 RelationshipId 精确获取 SlidePart
                    SlidePart slidePart = presPart.GetPartById(slideIdEntry.RelationshipId.Value) as SlidePart;
                    if (slidePart == null)
                    {
                        MessageBox.Show("无法加载幻灯片部分，文件可能已损坏。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 5. 在幻灯片中查找匹配的 Picture 元素并获取其关系 ID (rId)
                    string embedId = null;
                    var picture = slidePart.Slide
                        .Descendants<DocumentFormat.OpenXml.Presentation.Picture>()
                        .FirstOrDefault(p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value == shapeId);

                    if (picture != null)
                    {
                        embedId = picture.BlipFill?.Blip?.Embed?.Value;
                    }

                    if (string.IsNullOrEmpty(embedId))
                    {
                        MessageBox.Show("无法找到选中图片的内部引用关系，导出失败。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 6. 通过关系ID找到对应的 ImagePart
                    ImagePart imagePart = slidePart.GetPartById(embedId) as ImagePart;
                    if (imagePart == null)
                    {
                        MessageBox.Show("无法在文档包中找到图片数据，导出失败。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // 7. 准备保存文件对话框并导出
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        string originalFileName = Path.GetFileName(imagePart.Uri.OriginalString);
                        saveFileDialog.FileName = originalFileName;
                        saveFileDialog.Filter = "所有文件 (*.*)|*.*|PNG 图片 (*.png)|*.png|JPEG 图片 (*.jpg;*.jpeg)|*.jpg;*.jpeg|GIF 图片 (*.gif)|*.gif";
                        saveFileDialog.Title = "导出原始图片";

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            using (Stream imageStream = imagePart.GetStream())
                            using (FileStream fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create))
                            {
                                imageStream.CopyTo(fileStream);
                            }
                            MessageBox.Show($"图片已成功导出到：\n{saveFileDialog.FileName}", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"导出过程中发生错误：\n{ex.Message}", "意外错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void exportImageButton_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Presentation activePresentation;
            PowerPoint.Slides slides;

            try
            {
                activePresentation = pptApp.ActivePresentation;
                if (activePresentation == null)
                {
                    MessageBox.Show(
                        "没有打开的演示文稿可供导出。",
                        "无演示文稿",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }
                slides = activePresentation.Slides;
            }
            catch (Exception)
            {
                MessageBox.Show(
                    "无法访问演示文稿。请确保已打开一个演示文稿。",
                    "访问错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            if (slides == null || slides.Count == 0)
            {
                MessageBox.Show(
                    "当前演示文稿没有幻灯片。",
                    "无幻灯片",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            // Create export options dialog
            using (Form exportDialog = new Form())
            {
                exportDialog.Text = "导出设置";
                exportDialog.Width = 400;
                exportDialog.Height = 380; // Increase height for new PDF options and checkbox
                exportDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                exportDialog.StartPosition = FormStartPosition.CenterScreen;
                exportDialog.MaximizeBox = false;
                exportDialog.MinimizeBox = false;

                GroupBox rangeGroup = new GroupBox
                {
                    Text = "导出范围",
                    Location = new System.Drawing.Point(20, 20),
                    Width = 340,
                    Height = 85,
                };
                RadioButton currentSlideRadio = new RadioButton
                {
                    Text = "当前页",
                    Location = new System.Drawing.Point(20, 25),
                    Checked = true,
                    AutoSize = true,
                };
                RadioButton selectedSlidesRadio = new RadioButton
                {
                    Text = "选中的页面",
                    Location = new System.Drawing.Point(120, 25),
                    AutoSize = true,
                };
                RadioButton allSlidesRadio = new RadioButton
                {
                    Text = "全部页面",
                    Location = new System.Drawing.Point(20, 50),
                    AutoSize = true,
                };
                rangeGroup.Controls.AddRange(
                    new Control[] { currentSlideRadio, selectedSlidesRadio, allSlidesRadio }
                );

                GroupBox formatGroup = new GroupBox
                {
                    Text = "图片格式",
                    Location = new System.Drawing.Point(20, 115),
                    Width = 340,
                    Height = 60,
                };
                ComboBox formatCombo = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Location = new System.Drawing.Point(20, 25),
                    Width = 300,
                };
                formatCombo.Items.AddRange(new string[] { "PNG", "JPG", "BMP", "PDF" });
                formatCombo.SelectedIndex = 0; // Default to PNG
                formatGroup.Controls.Add(formatCombo);

                GroupBox pdfOptionsGroup = new GroupBox
                {
                    Text = "PDF 选项",
                    Location = new System.Drawing.Point(20, 185),
                    Width = 340,
                    Height = 60,
                    Visible = false,
                };
                CheckBox pdfSeparateFilesCheckBox = new CheckBox
                {
                    Text = "每页导出为单独的PDF文件",
                    Location = new System.Drawing.Point(20, 25),
                    AutoSize = true,
                    Checked = false,
                };
                pdfOptionsGroup.Controls.Add(pdfSeparateFilesCheckBox);

                GroupBox dpiGroup = new GroupBox
                {
                    Text = "导出DPI",
                    Location = new System.Drawing.Point(20, 185),
                    Width = 340,
                    Height = 60,
                }; // Adjusted Y
                ComboBox dpiCombo = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Location = new System.Drawing.Point(20, 25),
                    Width = 300,
                    BackColor = Color.White,
                };
                dpiCombo.Items.AddRange(new string[] { "96", "150", "300", "600" });
                dpiCombo.SelectedIndex = 2; // Default to 300 DPI
                dpiGroup.Controls.Add(dpiCombo);

                // Event handler for format change to hide/show DPI group and PDF options
                formatCombo.SelectedIndexChanged += (s, args) =>
                {
                    bool isPdf = formatCombo.SelectedItem.ToString().ToUpper() == "PDF";
                    dpiGroup.Visible = !isPdf;
                    pdfOptionsGroup.Visible = isPdf;
                    // Adjust layout if PDF options are shown/hidden
                    if (isPdf)
                    {
                        dpiGroup.Location = new System.Drawing.Point(
                            20,
                            185 + pdfOptionsGroup.Height + 10
                        ); // Move DPI group below PDF options
                    }
                    else
                    {
                        dpiGroup.Location = new System.Drawing.Point(20, 185); // Reset DPI group position
                    }
                };
                // Initial state for DPI group and PDF options visibility
                bool initialIsPdf = formatCombo.SelectedItem.ToString().ToUpper() == "PDF";
                dpiGroup.Visible = !initialIsPdf;
                pdfOptionsGroup.Visible = initialIsPdf;
                if (initialIsPdf)
                {
                    dpiGroup.Location = new System.Drawing.Point(
                        20,
                        185 + pdfOptionsGroup.Height + 10
                    );
                }

                // 添加导出后打开文件夹的复选框
                CheckBox openFolderCheckBox = new CheckBox
                {
                    Text = "导出完成后打开文件夹",
                    Location = new System.Drawing.Point(20, 255), // Adjusted Y position
                    AutoSize = true,
                    Checked = true, // 默认选中
                };

                Button okButton = new Button
                {
                    Text = "确定",
                    DialogResult = DialogResult.OK,
                    Location = new System.Drawing.Point(180, 285),
                    Width = 80,
                    Height = 36,
                }; // Adjusted Y
                Button cancelButton = new Button
                {
                    Text = "取消",
                    DialogResult = DialogResult.Cancel,
                    Location = new System.Drawing.Point(280, 285),
                    Width = 80,
                    Height = 36,
                }; // Adjusted Y

                exportDialog.Controls.AddRange(
                    new Control[]
                    {
                        rangeGroup,
                        formatGroup,
                        pdfOptionsGroup,
                        dpiGroup,
                        openFolderCheckBox,
                        okButton,
                        cancelButton,
                    }
                );
                exportDialog.AcceptButton = okButton;
                exportDialog.CancelButton = cancelButton;

                if (exportDialog.ShowDialog() == DialogResult.OK)
                {
                    string basePresentationName = "未命名";
                    string presentationCurrentFullPath = ""; // Best guess for the full path of the PPT file itself
                    string saveTargetDirectory;

                    try
                    {
                        string pptPathProperty = activePresentation.Path; // Can be URL or local directory path
                        string pptFullNameProperty = activePresentation.FullName; // Can be URL or local full file path

                        if (string.IsNullOrEmpty(pptPathProperty)) // Unsaved presentation
                        {
                            basePresentationName = "未命名";
                            presentationCurrentFullPath = Path.Combine(
                                Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
                                basePresentationName + ".pptx"
                            ); // Nominal path
                        }
                        else
                        {
                            basePresentationName = Path.GetFileNameWithoutExtension(
                                pptFullNameProperty
                            );
                            if (
                                string.IsNullOrEmpty(basePresentationName)
                                && !string.IsNullOrEmpty(pptPathProperty)
                            ) // Handle cases where FullName might be just a path
                            {
                                basePresentationName = Path.GetFileNameWithoutExtension(
                                    pptPathProperty
                                );
                            }
                            if (string.IsNullOrEmpty(basePresentationName))
                                basePresentationName = "未命名";

                            if (
                                pptPathProperty.StartsWith(
                                    "https://d.docs.live.net/",
                                    StringComparison.OrdinalIgnoreCase
                                )
                            )
                            {
                                string oneDriveRoot = GetLocalOneDrivePath();
                                if (
                                    !string.IsNullOrEmpty(oneDriveRoot)
                                    && Directory.Exists(oneDriveRoot)
                                )
                                {
                                    // Example URL: https://d.docs.live.net/USERID/Documents/MyPresentation.pptx
                                    // pathSegments: ["https:", "", "d.docs.live.net", "USERID", "Documents", "MyPresentation.pptx"] (from Split)
                                    string[] pathSegments = pptPathProperty.Split(
                                        new[] { '/' },
                                        StringSplitOptions.RemoveEmptyEntries
                                    );
                                    if (pathSegments.Length > 3) // Check for at least "https:", "d.docs.live.net", "USERID", and one more part
                                    {
                                        string relativePath = string.Join(
                                            Path.DirectorySeparatorChar.ToString(),
                                            pathSegments.Skip(3)
                                        );
                                        // For OneDrive URLs, construct local path by combining OneDrive root with the relative path
                                        // relativePath already includes the filename since we took all path segments after the user ID
                                        presentationCurrentFullPath = Path.Combine(
                                            oneDriveRoot,
                                            relativePath
                                        );
                                        // The line below was causing an extra folder with pptx name, pptFullNameProperty should be used for extension
                                        // presentationCurrentFullPath = Path.Combine(presentationCurrentFullPath, basePresentationName + ".pptx");
                                        // Corrected: pptFullNameProperty might already be the full local path if synced
                                        if (!File.Exists(Path.Combine(oneDriveRoot, relativePath))) // If relative path isn't the full file path
                                        {
                                            presentationCurrentFullPath = Path.Combine(
                                                oneDriveRoot,
                                                relativePath,
                                                Path.GetFileName(pptFullNameProperty)
                                            );
                                        }
                                        else
                                        {
                                            presentationCurrentFullPath = Path.Combine(
                                                oneDriveRoot,
                                                relativePath
                                            );
                                        }
                                    }
                                    else
                                    {
                                        // Fallback if URL structure is unexpected, try to use OneDrive root + filename
                                        presentationCurrentFullPath = Path.Combine(
                                            oneDriveRoot,
                                            basePresentationName
                                                + Path.GetExtension(pptFullNameProperty)
                                        );
                                    }
                                }
                                else
                                {
                                    presentationCurrentFullPath = Path.Combine(
                                        Environment.GetFolderPath(
                                            Environment.SpecialFolder.MyPictures
                                        ),
                                        basePresentationName
                                            + Path.GetExtension(pptFullNameProperty)
                                    );
                                }
                            }
                            else if (File.Exists(pptFullNameProperty)) // Local file or synced cloud file where FullName is a disk path
                            {
                                presentationCurrentFullPath = pptFullNameProperty;
                            }
                            else // Other web URLs or unresolvable local paths
                            {
                                string fallbackDir = GetLocalOneDrivePath();
                                if (
                                    string.IsNullOrEmpty(fallbackDir)
                                    || !Directory.Exists(fallbackDir)
                                )
                                {
                                    fallbackDir = Environment.GetFolderPath(
                                        Environment.SpecialFolder.MyPictures
                                    );
                                }
                                presentationCurrentFullPath = Path.Combine(
                                    fallbackDir,
                                    basePresentationName + Path.GetExtension(pptFullNameProperty)
                                );
                            }
                        }

                        // Determine the directory for SaveFileDialog: [DirectoryOfPPT]/[BasePresentationName]/
                        saveTargetDirectory = Path.GetDirectoryName(presentationCurrentFullPath);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"Error determining path information: {ex.Message}"
                        );
                        basePresentationName = "未命名";
                        saveTargetDirectory = Environment.GetFolderPath(
                            Environment.SpecialFolder.MyPictures
                        );
                    }

                    using (var saveDialog = new SaveFileDialog())
                    {
                        string selectedFormat = formatCombo.SelectedItem.ToString().ToUpper();
                        bool exportPdfAsSeparateFiles =
                            selectedFormat == "PDF" && pdfSeparateFilesCheckBox.Checked;

                        saveDialog.Filter = $"{selectedFormat} 文件|*.{selectedFormat.ToLower()}";
                        saveDialog.InitialDirectory = saveTargetDirectory;

                        bool exportCurrentSlide = currentSlideRadio.Checked;
                        bool exportSelectedSlides = selectedSlidesRadio.Checked;
                        PowerPoint.Slide slideToExport = null;
                        PowerPoint.SlideRange selectedSlideRange = null;

                        if (exportCurrentSlide)
                        {
                            try
                            {
                                slideToExport = pptApp.ActiveWindow.View.Slide;
                                saveDialog.FileName =
                                    $"{basePresentationName}_页面{slideToExport.SlideIndex}";
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    $"Error getting current slide: {ex.Message}"
                                );
                                MessageBox.Show(
                                    "无法获取当前幻灯片信息。将默认文件名。",
                                    "警告",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning
                                );
                                saveDialog.FileName = $"{basePresentationName}_页面";
                            }
                        }
                        else if (exportSelectedSlides)
                        {
                            try
                            {
                                if (
                                    pptApp.ActiveWindow.Selection.Type
                                    == PpSelectionType.ppSelectionSlides
                                )
                                {
                                    selectedSlideRange = pptApp.ActiveWindow.Selection.SlideRange;
                                    if (selectedSlideRange.Count > 0)
                                    {
                                        if (
                                            selectedFormat == "PDF" && !exportPdfAsSeparateFiles
                                            || selectedSlideRange.Count == 1
                                        )
                                        {
                                            saveDialog.FileName = $"{basePresentationName}_页面"; // For single PDF or single selected slide
                                        }
                                        else
                                        {
                                            // For multiple slides to image format, or separate PDFs, suggest a base name
                                            saveDialog.FileName = $"{basePresentationName}_页面";
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show(
                                            "没有选中的幻灯片。请先选择幻灯片。",
                                            "无选中幻灯片",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Warning
                                        );
                                        return;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(
                                        "请先在幻灯片浏览视图或大纲视图中选择幻灯片。",
                                        "选择模式错误",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Warning
                                    );
                                    return;
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    $"Error getting selected slides: {ex.Message}"
                                );
                                MessageBox.Show(
                                    "无法获取选中的幻灯片信息。将默认文件名。",
                                    "警告",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning
                                );
                                saveDialog.FileName = $"{basePresentationName}_选中的页面";
                            }
                        }
                        else // Exporting all slides
                        {
                            if (selectedFormat == "PDF" && !exportPdfAsSeparateFiles)
                            {
                                saveDialog.FileName = $"{basePresentationName}";
                            }
                            else
                            {
                                saveDialog.FileName = $"{basePresentationName}_页面";
                            }
                        }

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            string exportPath = saveDialog.FileName; // Full path from SaveFileDialog
                            string exportDirectory = Path.GetDirectoryName(exportPath);
                            string baseExportFileName = Path.GetFileNameWithoutExtension(
                                exportPath
                            );

                            // Ensure the target directory exists before exporting
                            if (!Directory.Exists(exportDirectory))
                            {
                                try
                                {
                                    Directory.CreateDirectory(exportDirectory);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(
                                        $"无法创建导出目录 '{exportDirectory}': {ex.Message}",
                                        "目录创建错误",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error
                                    );
                                    return;
                                }
                            }

                            try
                            {
                                if (selectedFormat == "PDF")
                                {
                                    if (exportPdfAsSeparateFiles)
                                    {
                                        if (exportCurrentSlide && slideToExport != null)
                                        {
                                            string filePath = Path.Combine(
                                                exportDirectory,
                                                $"{baseExportFileName}.pdf"
                                            );
                                            ExportSlideAsPdf(
                                                activePresentation,
                                                slideToExport.SlideIndex,
                                                filePath
                                            );
                                        }
                                        else if (
                                            exportSelectedSlides
                                            && selectedSlideRange != null
                                            && selectedSlideRange.Count > 0
                                        )
                                        {
                                            foreach (PowerPoint.Slide slide in selectedSlideRange)
                                            {
                                                string filePath = Path.Combine(
                                                    exportDirectory,
                                                    $"{baseExportFileName}{slide.SlideIndex}.pdf"
                                                );
                                                ExportSlideAsPdf(
                                                    activePresentation,
                                                    slide.SlideIndex,
                                                    filePath
                                                );
                                            }
                                        }
                                        else if (!exportCurrentSlide && !exportSelectedSlides) // All slides
                                        {
                                            for (int i = 1; i <= slides.Count; i++)
                                            {
                                                PowerPoint.Slide slide = slides[i];
                                                string filePath = Path.Combine(
                                                    exportDirectory,
                                                    $"{baseExportFileName}{slide.SlideIndex}.pdf"
                                                );
                                                ExportSlideAsPdf(
                                                    activePresentation,
                                                    slide.SlideIndex,
                                                    filePath
                                                );
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(
                                                    slide
                                                );
                                                slide = null;
                                            }
                                        }
                                    }
                                    else // Export as a single PDF
                                    {
                                        if (exportCurrentSlide && slideToExport != null)
                                        {
                                            activePresentation
                                                .Slides.Range(
                                                    new int[] { slideToExport.SlideIndex }
                                                )
                                                .Select();
                                            activePresentation.ExportAsFixedFormat(
                                                Path: exportPath,
                                                FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF,
                                                Intent: PpFixedFormatIntent.ppFixedFormatIntentPrint,
                                                OutputType: PpPrintOutputType.ppPrintOutputSlides,
                                                RangeType: PpPrintRangeType.ppPrintSelection
                                            );
                                        }
                                        else if (
                                            exportSelectedSlides
                                            && selectedSlideRange != null
                                            && selectedSlideRange.Count > 0
                                        )
                                        {
                                            // For PDF export of selected slides, PowerPoint handles this via selection
                                            // Ensure the slides are actually selected in the UI for ExportAsFixedFormat to work correctly with ppPrintSelection
                                            // It's generally better to rely on the user having them selected,
                                            // but programmatically selecting them can be an option if needed, though it might change user's view.
                                            // For simplicity, we assume they are already selected as per the radio button choice.
                                            // If direct API for exporting a SlideRange to PDF existed, it would be cleaner.
                                            // The most robust way for selected slides to PDF is to ensure they are selected, then use ppPrintSelection.
                                            // PowerPoint's UI "Save As PDF" with "Options..." -> "Selection" does this.
                                            // We will select them programmatically before export.
                                            int[] slideIndices = new int[selectedSlideRange.Count];
                                            for (int i = 0; i < selectedSlideRange.Count; i++)
                                            {
                                                slideIndices[i] = selectedSlideRange[
                                                    i + 1
                                                ].SlideIndex;
                                            }
                                            activePresentation.Slides.Range(slideIndices).Select();

                                            activePresentation.ExportAsFixedFormat(
                                                Path: exportPath,
                                                FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF,
                                                Intent: PpFixedFormatIntent.ppFixedFormatIntentPrint,
                                                OutputType: PpPrintOutputType.ppPrintOutputSlides,
                                                RangeType: PpPrintRangeType.ppPrintSelection // Export only the selected slides
                                            );
                                        }
                                        else if (!exportCurrentSlide && !exportSelectedSlides) // All slides
                                        {
                                            activePresentation.ExportAsFixedFormat(
                                                Path: exportPath,
                                                FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF,
                                                Intent: PpFixedFormatIntent.ppFixedFormatIntentPrint
                                            ); // Defaults to all slides
                                        }
                                        else if (exportCurrentSlide && slideToExport == null)
                                        {
                                            MessageBox.Show(
                                                "无法导出当前幻灯片为PDF，因为它未被正确识别。",
                                                "导出错误",
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error
                                            );
                                            return;
                                        }
                                    }
                                }
                                else // Image formats
                                {
                                    int dpi = int.Parse(dpiCombo.SelectedItem.ToString());
                                    if (exportCurrentSlide && slideToExport != null)
                                    {
                                        ExportSlide(slideToExport, exportPath, selectedFormat, dpi);
                                    }
                                    else if (
                                        exportSelectedSlides
                                        && selectedSlideRange != null
                                        && selectedSlideRange.Count > 0
                                    )
                                    {
                                        string outputFileNameBase =
                                            Path.GetFileNameWithoutExtension(exportPath); // Base name from SaveDialog
                                        if (selectedSlideRange.Count == 1)
                                        {
                                            ExportSlide(
                                                selectedSlideRange[1],
                                                exportPath,
                                                selectedFormat,
                                                dpi
                                            );
                                        }
                                        else
                                        {
                                            for (int i = 1; i <= selectedSlideRange.Count; i++)
                                            {
                                                PowerPoint.Slide slide = selectedSlideRange[i];
                                                // Use the slide's actual index for a more consistent naming if desired, or just a sequence number
                                                string filename = Path.Combine(
                                                    exportDirectory,
                                                    $"{outputFileNameBase}{slide.SlideIndex}.{selectedFormat.ToLower()}"
                                                );
                                                ExportSlide(slide, filename, selectedFormat, dpi);
                                                // No need to release com object for slide from SlideRange here as it's managed by the range
                                            }
                                        }
                                    }
                                    else if (!exportCurrentSlide && !exportSelectedSlides) // All slides
                                    {
                                        string outputFileNameBase =
                                            Path.GetFileNameWithoutExtension(exportPath); // Base name from SaveDialog
                                        for (int i = 1; i <= slides.Count; i++)
                                        {
                                            PowerPoint.Slide slide = slides[i];
                                            string filename = Path.Combine(
                                                exportDirectory,
                                                $"{outputFileNameBase}{slide.SlideIndex}.{selectedFormat.ToLower()}"
                                            );
                                            ExportSlide(slide, filename, selectedFormat, dpi);
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(
                                                slide
                                            );
                                            slide = null;
                                        }
                                    }
                                    else if (exportCurrentSlide && slideToExport == null)
                                    {
                                        MessageBox.Show(
                                            "无法导出当前幻灯片，因为它未被正确识别。",
                                            "导出错误",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Error
                                        );
                                        return;
                                    }
                                }

                                MessageBox.Show(
                                    "导出完成！",
                                    "成功",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                                // 根据复选框状态决定是否打开文件夹
                                if (openFolderCheckBox.Checked)
                                {
                                    System.Diagnostics.Process.Start(
                                        "explorer.exe",
                                        exportDirectory
                                    );
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(
                                    $"导出过程中发生错误：{ex.Message}",
                                    "错误",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                            }
                            finally
                            {
                                if (slideToExport != null)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(
                                        slideToExport
                                    );
                                if (selectedSlideRange != null)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(
                                        selectedSlideRange
                                    );
                            }
                        }
                    }
                }
            }
        }

        private void ExportSlideAsPdf(
            PowerPoint.Presentation presentation,
            int slideIndex,
            string filePath
        )
        {
            presentation.ExportAsFixedFormat(
                Path: filePath,
                FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF,
                Intent: PpFixedFormatIntent.ppFixedFormatIntentPrint,
                PrintRange: presentation.PrintOptions.Ranges.Add(slideIndex, slideIndex) // Export specific slide
            );
            // Clean up the added print range to avoid issues with subsequent exports
            if (presentation.PrintOptions.Ranges.Count > 0)
            {
                // PowerPoint's PrintOptions.Ranges collection is 1-based.
                // And it seems it might accumulate ranges if not cleared.
                // A robust way is to clear all ranges after use if they are not meant to be persistent.
                // However, directly clearing all might affect other print settings if the user configured them.
                // For this specific export, we add a range, use it, and ideally, it should be self-contained.
                // If issues arise, clearing might be needed:
                // while (presentation.PrintOptions.Ranges.Count > 0) {
                //     presentation.PrintOptions.Ranges[1].Delete();
                // }
                // For now, assume PowerPoint handles the temporary range correctly for ExportAsFixedFormat.
                // If exporting multiple single-slide PDFs in a loop, ensure ranges are managed.
                // A safer approach for single slide export is to select it and use ppPrintSelection.
                // However, the PrintRange approach is more direct if it works reliably across versions.

                // Let's try selecting the slide and using ppPrintSelection for single slide PDF export
                // This is generally more reliable.
                presentation.Slides.Range(new int[] { slideIndex }).Select();
                presentation.ExportAsFixedFormat(
                    Path: filePath,
                    FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF,
                    Intent: PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    OutputType: PpPrintOutputType.ppPrintOutputSlides,
                    RangeType: PpPrintRangeType.ppPrintSelection
                );
            }
        }

        /// <summary>
        /// 复制选中图片的原始数据到剪贴板，以保证最高质量。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        // 放在你的 Ribbon 类或者 ThisAddIn 类中
        // 确保已经 using 了 System.Windows.Forms, System.IO, System.Drawing,
        // PowerPoint = Microsoft.Office.Interop.PowerPoint, Office = Microsoft.Office.Core

        private void CopyOriginalPicture_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection sel = null;
            PowerPoint.Shape selectedShape = null;

            try
            {
                if (app.ActiveWindow == null || app.ActiveWindow.View == null) return;
                sel = app.ActiveWindow.Selection;

                // 1. 验证是否选择了单个图片
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count != 1)
                {
                    MessageBox.Show("请选择单个图片对象。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                selectedShape = sel.ShapeRange[1];

                // if (selectedShape.Type != Office.MsoShapeType.msoPicture && selectedShape.Type != Office.MsoShapeType.msoLinkedPicture)
                // {
                //     MessageBox.Show("所选对象不是图片，请重新选择。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //     return;
                // }

                // 2. 保存图片的原始状态（尺寸、位置和锁定设置）
                float originalWidth = selectedShape.Width;
                float originalHeight = selectedShape.Height;
                float originalLeft = selectedShape.Left;
                float originalTop = selectedShape.Top;
                Office.MsoTriState originalLockAspectRatio = selectedShape.LockAspectRatio;

                try
                {
                    // 3. 获取幻灯片的高度
                    float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                    // 4. 临时修改图片尺寸
                    // 确保锁定宽高比，以便在调整高度时宽度能按比例缩放
                    selectedShape.LockAspectRatio = Office.MsoTriState.msoTrue;
                    // 将图片高度设置为幻灯片的高度
                    selectedShape.Height = slideHeight;

                    // 5. 直接将当前状态的形状复制到剪贴板
                    selectedShape.Copy();

                    //MessageBox.Show("已成功将图片（按幻灯片高度缩放后）复制到剪贴板！", "复制成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"复制图片时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // 6. **关键步骤**: 无论成功或失败，都恢复图片的原始状态
                    if (selectedShape != null)
                    {
                        try
                        {
                            // 按相反的顺序恢复，先恢复锁定状态，再恢复尺寸和位置
                            selectedShape.LockAspectRatio = originalLockAspectRatio;
                            selectedShape.Width = originalWidth;
                            selectedShape.Height = originalHeight;
                            selectedShape.Left = originalLeft;
                            selectedShape.Top = originalTop;
                        }
                        catch (Exception restoreEx)
                        {
                            // 如果恢复失败，在调试时输出信息，通常不打扰用户
                            System.Diagnostics.Debug.WriteLine($"恢复图片原始状态失败: {restoreEx.Message}");
                        }
                    }
                }
            }
            catch (Exception outerEx)
            {
                MessageBox.Show($"发生意外错误: {outerEx.Message}", "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 7. 释放COM对象
                if (selectedShape != null) Marshal.ReleaseComObject(selectedShape);
                if (sel != null) Marshal.ReleaseComObject(sel);
            }
        }

        private void ExportSlide(PowerPoint.Slide slide, string filename, string format, int dpi)
        {
            string upperFormat = format.ToUpper(); // Ensure consistent case for comparison
            if (upperFormat == "SVG")
            {
                slide.Export(filename, "SVG");
            }
            else
            {
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;

                // 计算导出尺寸
                int exportWidth = (int)((slideWidth / 72.0f) * dpi);
                int exportHeight = (int)((slideHeight / 72.0f) * dpi);

                slide.Export(filename, format, exportWidth, exportHeight);
            }
        }


        // Helper method to get the local OneDrive path
        private string GetLocalOneDrivePath()
        {
            // Try environment variables first
            string oneDrivePath = Environment.GetEnvironmentVariable("OneDrive");
            if (!string.IsNullOrEmpty(oneDrivePath) && Directory.Exists(oneDrivePath))
            {
                return oneDrivePath;
            }

            // Try consumer OneDrive path
            oneDrivePath = Environment.GetEnvironmentVariable("OneDriveConsumer");
            if (!string.IsNullOrEmpty(oneDrivePath) && Directory.Exists(oneDrivePath))
            {
                return oneDrivePath;
            }

            // Try business OneDrive path
            oneDrivePath = Environment.GetEnvironmentVariable("OneDriveCommercial");
            if (!string.IsNullOrEmpty(oneDrivePath) && Directory.Exists(oneDrivePath))
            {
                return oneDrivePath;
            }

            // Fallback to registry lookup for OneDrive path
            try
            {
                using (
                    var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\OneDrive"
                    )
                )
                {
                    if (key != null)
                    {
                        oneDrivePath = key.GetValue("UserFolder") as string;
                        if (!string.IsNullOrEmpty(oneDrivePath) && Directory.Exists(oneDrivePath))
                        {
                            return oneDrivePath;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Error accessing registry for OneDrive path: {ex.Message}"
                );
            }

            return null;
        }
        private void insertLatexSVG_Click(object sender, RibbonControlEventArgs e)
        {
            if (app == null)
            {
                app = Globals.ThisAddIn.Application;
            }

            if (app?.ActiveWindow == null || app.ActiveWindow.View == null)
            {
                MessageBox.Show("当前没有打开的演示文稿窗口。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Slide currentSlide = null;
            try
            {
                currentSlide = app.ActiveWindow.View.Slide;
            }
            catch (Exception)
            {
                // 如果无法获取当前幻灯片，则提示并返回
            }

            if (currentSlide == null)
            {
                MessageBox.Show("无法获取当前幻灯片，请确保已经打开并选中了幻灯片。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (Form inputDialog = new Form
            {
                Width = 520,
                Height = 420,
                Text = "输入 LaTeX 公式",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
            })
            {
                var infoLabel = new Label
                {
                    Text = "请输入 LaTeX 数学公式，不需要输入 $ 或 \\(...\\)。",
                    Dock = DockStyle.Top,
                    Height = 40,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Padding = new Padding(10, 10, 10, 0),
                };

                var latexInputBox = new TextBox
                {
                    Multiline = true,
                    Dock = DockStyle.Fill,
                    Font = new Font("Consolas", 12),
                    ScrollBars = ScrollBars.Vertical,
                };

                var buttonPanel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 50,
                    Padding = new Padding(10),
                    FlowDirection = FlowDirection.RightToLeft,
                    AutoSize = false,
                };

                var okButton = new Button
                {
                    Text = "确定",
                    DialogResult = DialogResult.OK,
                    Width = 80,
                    Height = 30,
                    Margin = new Padding(10, 10, 0, 0),
                };

                var cancelButton = new Button
                {
                    Text = "取消",
                    DialogResult = DialogResult.Cancel,
                    Width = 80,
                    Height = 30,
                    Margin = new Padding(10, 10, 0, 0),
                };

                buttonPanel.Controls.Add(okButton);
                buttonPanel.Controls.Add(cancelButton);

                inputDialog.Controls.Add(latexInputBox);
                inputDialog.Controls.Add(buttonPanel);
                inputDialog.Controls.Add(infoLabel);

                inputDialog.AcceptButton = okButton;
                inputDialog.CancelButton = cancelButton;

                if (inputDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                string latexInput = latexInputBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(latexInput))
                {
                    return;
                }

                string normalizedLatex = NormalizeLatexInput(latexInput);

                try
                {
                    var converter = Globals.ThisAddIn.LatexSvgConverter ?? new LatexToSvgConverter();
                    string svgContent = converter.ConvertLatexToSvg(normalizedLatex);

                    if (string.IsNullOrWhiteSpace(svgContent))
                    {
                        MessageBox.Show("未生成有效的 SVG 内容。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    string tempDirectory = Path.Combine(Path.GetTempPath(), "SlideSCI", "latex");
                    Directory.CreateDirectory(tempDirectory);
                    string tempFilePath = Path.Combine(tempDirectory, $"latex_{Guid.NewGuid():N}.svg");
                    File.WriteAllText(tempFilePath, svgContent, Encoding.UTF8);

                    try
                    {
                        Shape svgShape = currentSlide.Shapes.AddPicture(
                            tempFilePath,
                            Office.MsoTriState.msoFalse,
                            Office.MsoTriState.msoTrue,
                            0,
                            0,
                            -1,
                            -1
                        );

                        if (svgShape != null)
                        {
                            svgShape.Left = (currentSlide.Master.Width - svgShape.Width) / 2;
                            svgShape.Top = (currentSlide.Master.Height - svgShape.Height) / 2;
                            svgShape.LockAspectRatio = Office.MsoTriState.msoTrue;
                            svgShape.AlternativeText = $"{normalizedLatex}";
                            // svgShape 默认设置为2x大小
                            svgShape.Width *= 2;
                            svgShape.Select();
                        }
                    }
                    finally
                    {
                        try
                        {
                            if (File.Exists(tempFilePath))
                            {
                                File.Delete(tempFilePath);
                            }
                        }
                        catch (Exception deleteEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"删除临时 SVG 文件失败: {deleteEx.Message}");
                        }
                    }
                }
                catch (FileNotFoundException fnfEx)
                {
                    MessageBox.Show(
                        $"{fnfEx.Message}\n请在插件目录下的 latex-converter 文件夹中执行 npm install。",
                        "脚本缺失",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
                catch (InvalidOperationException invalidOpEx)
                {
                    MessageBox.Show(
                        invalidOpEx.Message,
                        "LaTeX 转换失败",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"插入 LaTeX SVG 时发生错误: {ex.Message}",
                        "错误",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }

        private static string NormalizeLatexInput(string latexInput)
        {
            string trimmed = latexInput.Trim();

            if (trimmed.StartsWith("$$") && trimmed.EndsWith("$$"))
            {
                trimmed = trimmed.Substring(2, trimmed.Length - 4);
            }
            else if (trimmed.StartsWith("$") && trimmed.EndsWith("$"))
            {
                trimmed = trimmed.Substring(1, trimmed.Length - 2);
            }
            else if (trimmed.StartsWith(@"\(") && trimmed.EndsWith(@"\)"))
            {
                trimmed = trimmed.Substring(2, trimmed.Length - 4);
            }
            else if (trimmed.StartsWith(@"\[") && trimmed.EndsWith(@"\]"))
            {
                trimmed = trimmed.Substring(2, trimmed.Length - 4);
            }

            return trimmed.Replace("\r", "").Trim();
        }
        private void updateLabelsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
            {
                MessageBox.Show("请选择要更新标签的文本框。");
                return;
            }

            string fontFamily = labelFontNameEditBox.Text;
            float fontSize;
            if (!float.TryParse(labelFontSizeEditBox.Text, out fontSize))
            {
                MessageBox.Show("请输入有效的字体大小。");
                return;
            }

            string labelTemplate = labelTemplateComboBox.Text;

            var templates = new Dictionary<string, string>
            {
                { "A", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" },
                { "a", "abcdefghijklmnopqrstuvwxyz" },
                { "A)", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" },
                { "a)", "abcdefghijklmnopqrstuvwxyz" },
                { "(A)", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" },
                { "(a)", "abcdefghijklmnopqrstuvwxyz" },
                { "1", "123456789" }, 
                { "1)", "123456789" }, 
                { "Ⅰ", "ⅠⅡⅢⅣⅤⅥⅦⅦⅨⅩ" },
                { "Ⅰ)", "ⅠⅡⅢⅣⅤⅥⅦⅦⅨⅩ" },
                { "①", "①②③④⑤⑥⑦⑧⑨⑩" },
                { "①)", "①②③④⑤⑥⑦⑧⑨⑩" },
                { "一", "一二三四五六七八九十" },
                { "一)", "一二三四五六七八九十" },
            };

            if (!templates.ContainsKey(labelTemplate))
            {
                labelTemplate = "A";
            }

            string labels = templates[labelTemplate];
            bool isNumeric = labelTemplate.StartsWith("1");

            // 获取起始编号
            int startIndex = 1;
            if (!string.IsNullOrEmpty(labelIndex.Text) && !int.TryParse(labelIndex.Text, out startIndex))
            {
                MessageBox.Show("请输入有效的起始编号。");
                return;
            }

            // 过滤出文本框
            var textBoxes = new List<Shape>();
            foreach (Shape shape in sel.ShapeRange)
            {
                if (shape.Type == Office.MsoShapeType.msoTextBox)
                {
                    textBoxes.Add(shape);
                }
            }

            if (textBoxes.Count == 0)
            {
                MessageBox.Show("请选择要更新标签的文本框。");
                return;
            }

            // 对文本框进行排序（按位置）
            var groups = new List<ImageGroup>();
            foreach (var textBox in textBoxes)
            {
                bool addedToExistingGroup = false;
                foreach (var group in groups)
                {
                    if (group.OverlapsWith(textBox))
                    {
                        group.AddShape(textBox);
                        addedToExistingGroup = true;
                        break;
                    }
                }

                if (!addedToExistingGroup)
                {
                    var newGroup = new ImageGroup();
                    newGroup.AddShape(textBox);
                    groups.Add(newGroup);
                }
            }

            // Sort shapes within each group by x position
            foreach (var group in groups)
            {
                group.Shapes.Sort((a, b) => a.Left.CompareTo(b.Left));
            }

            // Sort groups by MinTop
            groups.Sort((a, b) => a.MinTop.CompareTo(b.MinTop));

            // Create flattened list of sorted shapes
            var sortedTextBoxes = new List<Shape>();
            foreach (var group in groups)
            {
                sortedTextBoxes.AddRange(group.Shapes);
            }

            // 更新文本框的标签
            for (int i = 0; i < sortedTextBoxes.Count; i++)
            {
                try
                {
                    var textBox = sortedTextBoxes[i];
                    string label;
                    if (isNumeric)
                    {
                        label = (startIndex + i).ToString();
                    }
                    else
                    {
                        int labelIndexValue = (startIndex - 1 + i) % labels.Length;
                        label = labels[labelIndexValue].ToString();
                    }

                    if (labelTemplate.EndsWith(")") && !labelTemplate.StartsWith("("))
                    {
                        label += ")";
                    }
                    else if (labelTemplate.StartsWith("(") && labelTemplate.EndsWith(")"))
                    {
                        label = "(" + label + ")";
                    }

                    // 更新文本框内容和字体设置
                    textBox.TextFrame.TextRange.Text = label;
                    textBox.TextFrame.TextRange.Font.Size = fontSize;
                    textBox.TextFrame.TextRange.Font.Name = fontFamily;
                    
                    // 应用加粗设置
                    if (labelBoldcheckBox.Checked)
                    {
                        textBox.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    }
                    else
                    {
                        textBox.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"更新标签时出错: {ex.Message}");
                }
            }

            // 如果开启了编号自动更新，则更新起始编号
            if (labelIndexUpdatecheckBox.Checked)
            {
                int nextIndex = startIndex + sortedTextBoxes.Count;
                labelIndex.Text = nextIndex.ToString();
            }
        }
    }
}

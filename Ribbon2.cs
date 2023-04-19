using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace IMAGER
{
    public partial class Ribbon2
    {
        PowerPoint.Application app;
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //获取页面目前的尺寸
            float w = app.ActivePresentation.PageSetup.SlideWidth;
            float h = app.ActivePresentation.PageSetup.SlideHeight;

            //让选中的图片铺满屏幕
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Selection sel = app.ActiveWindow.Selection; 
            PowerPoint.ShapeRange range = sel.ShapeRange;
            range[1].LockAspectRatio = Office.MsoTriState.msoFalse; 
            range[1].Width = w;
            range[1].Height = h; 
            range[1].Left = 0; 
            range[1].Top = 0;


            //插入一个全屏矩形，填充为黑色半透明
            PowerPoint.Shape shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 0, 0, w, h);
            Random rd = new Random(); 
            int r = rd.Next(0,30); 
            int g = rd.Next(10,50); 
            int b = rd.Next(20,70);
            shape.Fill.ForeColor.RGB = r + g * 256 + b * 256 * 256;
            shape.Fill.Transparency = 0.4F;
            shape.Line.Visible = Office.MsoTriState.msoFalse;

            //插入一个文本框
            PowerPoint.Shape txb = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, w / 4, h / 3, w / 2, h / 3);
            txb.TextFrame.TextRange.Text = "LAOHEI";
            txb.TextFrame2.TextRange.Font.Size = 96;
            txb.TextFrame2.TextRange.Font.NameFarEast = "微软雅黑"; 
            txb.TextFrame2.TextRange.Font.Name = "微软雅黑";
            txb.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215;
            txb.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter; 
            txb.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

        }
    }
}

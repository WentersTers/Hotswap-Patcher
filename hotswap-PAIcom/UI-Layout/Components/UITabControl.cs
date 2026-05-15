using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using PAIcomPatcher.UI.Styling;

namespace PAIcomPatcher.UI.Components;

public sealed class UITabControl : TabControl
{
    public Color TabBackColor { get; set; } = Color.FromArgb(45, 45, 48);
    public Color TabSelectedBackColor { get; set; } = Color.FromArgb(63, 63, 70);
    public Color TabTextColor { get; set; } = Color.White;

    public UITabControl()
    {
        DrawMode = TabDrawMode.OwnerDrawFixed;
        SizeMode = TabSizeMode.Fixed;
        ItemSize = new Size(140, 30);
        Padding = new Point(12, 6);
        Appearance = TabAppearance.Normal;
    }

    protected override void OnDrawItem(DrawItemEventArgs e)
    {
        var page = TabPages[e.Index];
        var tabRect = GetTabRect(e.Index);
        var selected = e.Index == SelectedIndex;
        var fillColor = selected ? TabSelectedBackColor : TabBackColor;

        using var fillBrush = new SolidBrush(fillColor);
        using var borderPen = new Pen(ControlPaint.Dark(fillColor));

        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.FillRectangle(fillBrush, tabRect);
        e.Graphics.DrawRectangle(borderPen, tabRect);

        var textRect = Rectangle.Inflate(tabRect, -8, -4);
        TextRenderer.DrawText(e.Graphics, page.Text, Font, textRect, TabTextColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    }
}
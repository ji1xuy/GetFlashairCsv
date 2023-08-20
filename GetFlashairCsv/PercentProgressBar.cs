//ProgressBarにパーセンテージなどの文字列を表示する
//https://dobon.net/vb/dotnet/control/pbshowtext.html

using System;
using System.Drawing;
using System.Windows.Forms;

/// <summary>
/// パーセント表示のあるプログレスバーコントロールを表します。
/// </summary>
public class PercentProgressBar : ProgressBar
{
    public PercentProgressBar()
    {
        //Paintイベントが発生するようにする
        //ダブルバッファリングを有効にする
        base.SetStyle(ControlStyles.UserPaint |
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer, true);
    }

    public string DisplayText { get; set; } = "";

    protected override void OnPaint(PaintEventArgs e)
    {
        //背景を描画する
        ProgressBarRenderer.DrawHorizontalBar(e.Graphics, this.ClientRectangle);

        //バーの幅を計算する
        //上下左右1ピクセルの枠の部分を除外して計算する
        double percent = (double)(this.Value - this.Minimum)
            / (double)(this.Maximum - this.Minimum);
        int chunksWidth = (int)((double)(this.ClientSize.Width - 2) * percent);
        System.Drawing.Rectangle chunksRect = new System.Drawing.Rectangle(1, 1,
            chunksWidth, this.ClientSize.Height - 2);
        //バーを描画する
        ProgressBarRenderer.DrawHorizontalChunks(e.Graphics, chunksRect);

        //表示する文字列を決定する
        //string displayText = string.Format("{0}%", (int)(percent * 100.0));
        //文字列を描画する
        TextFormatFlags tff = TextFormatFlags.HorizontalCenter |
            TextFormatFlags.VerticalCenter |
            TextFormatFlags.SingleLine;
        //TextRenderer.DrawText(e.Graphics, displayText, this.Font,
        //    this.ClientRectangle, SystemColors.ControlText, tff);
        TextRenderer.DrawText(e.Graphics, DisplayText, this.Font,
            this.ClientRectangle, SystemColors.ControlText, tff);
    }
}
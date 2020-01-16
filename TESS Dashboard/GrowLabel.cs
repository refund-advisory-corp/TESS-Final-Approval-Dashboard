using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;

public class GrowLabel : Label
{
    public int SetWidth { get; set; }

    private bool mGrowing;
    public GrowLabel()
    {
        this.AutoSize = false;
    }
    private void resizeLabel()
    {
        if (mGrowing) return;
        try
        {
            string Rem = this.Text;
            mGrowing = true;

            //Size sz = new Size(this.Width, Int32.MaxValue);

            int Progress = 0;
            //Keep fixing the string until it fits.
            while (TextRenderer.MeasureText(Rem, this.Font, new Size(Int32.MaxValue, Int32.MaxValue)).Width > SetWidth)
            {
                //Now, for a single fix, we add one newline where it needs to be. To find where we need to do this, keep testing different string lengths.
                int InstallLength = 1;

                while (TextRenderer.MeasureText(Rem.Substring(Progress, InstallLength), this.Font, new Size(Int32.MaxValue, Int32.MaxValue)).Width < SetWidth)
                {
                    InstallLength += 1;
                }

                Progress += InstallLength;
                Rem = Rem.Substring(0, Progress - 1) + System.Environment.NewLine + Rem.Substring(Progress - 1, Rem.Length - (Progress - 1));
                
            }

            //sz = TextRenderer.MeasureText(this.Text, this.Font, sz, TextFormatFlags.WordBreak);
            //this.Height = sz.Height;

            this.Text = Rem;
            this.Width = SetWidth;
            this.Height = TextRenderer.MeasureText(Rem, this.Font, new Size(this.Width, Int32.MaxValue)).Height;
        }
        finally
        {
            mGrowing = false;
        }
    }
    protected override void OnTextChanged(EventArgs e)
    {
        base.OnTextChanged(e);
        resizeLabel();
    }
    protected override void OnFontChanged(EventArgs e)
    {
        base.OnFontChanged(e);
        resizeLabel();
    }
    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        resizeLabel();
    }
}

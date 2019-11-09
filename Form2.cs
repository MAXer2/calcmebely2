
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
 
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        Bitmap RotatingBlocks;
        Point DrawHere;
        Rectangle InvalidRect;
        
        public Form1()
        {
           
            RotatingBlocks = new Bitmap("d:\\AG00011_.GIF");
            DrawHere = new Point(10, 10);
            InvalidRect = new Rectangle(DrawHere, RotatingBlocks.Size);
 
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.DoubleBuffer, true);
           
 
        }
 
       
        protected override void OnPaint(PaintEventArgs e)
        {                      
            ImageAnimator.UpdateFrames(RotatingBlocks);
            
            e.Graphics.DrawImage(RotatingBlocks, DrawHere);
        }
 
        private void OnFrameChanged(object o, EventArgs e)
        {
            
            this.Invalidate(InvalidRect);
        }
 
        private void Form1_Load(object sender, EventArgs e)
        {
            if (ImageAnimator.CanAnimate(RotatingBlocks))
            {
                
                ImageAnimator.Animate(RotatingBlocks,
                                      new EventHandler(this.OnFrameChanged));
                
               
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(526, 515);
            this.Name = "Form1";
            this.ResumeLayout(false);

        }
 
    }
}
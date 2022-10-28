using System;
using System.Drawing;
using System.Windows.Forms;

namespace Language_Learning_Winform_Entry_2
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

            int i = 0;
            int allButtonLocationX = 0;
            int allButtonLocationY = 0;
            switch (page)
            {
                case 1:
                    labelQuest.Location = new Point((int)(0.5 * (this.Width - labelQuest.Width)), (int)(0.1 * this.Height));
                    buttonChooseFile.Location = new Point((int)(0.5 * (this.Width - buttonChooseFile.Width)), (int)(0.6 * this.Height));
                    break;
                case 2:
                    labelQuest2.Location = new Point((int)(0.5 * (this.Width - labelQuest2.Width)), (int)(0.1 * this.Height));
                    panelModeChoose.Width = this.Width;
                    panelModeChoose.Height = (int)(this.Height - 150);
                    panelModeChoose.Location = new Point(0, (int)(0.15 * this.Height));

                    allButtonLocationX = 50;
                    allButtonLocationY = labelQuest2.Height;
                    foreach (Button btn in buttonModeChoose)
                    {
                        //btn.Location = new Point((int)(0.5 * (panelModeChoose.Width - btn.Width)), labelQuest2.Height + (int)(0.1 * (i + 1) * panelModeChoose.Height));
                        //i = i + 1;
                        if (allButtonLocationX + btn.Width + 50 > this.Width)
                        {
                            allButtonLocationX = 50;
                            allButtonLocationY += (int)(btn.Height + 30);
                            btn.Location = new Point(allButtonLocationX, allButtonLocationY);
                            allButtonLocationX += (btn.Width + 50);
                        }
                        else
                        {
                            btn.Location = new Point(allButtonLocationX, allButtonLocationY);
                            allButtonLocationX += (btn.Width + 50);
                        }

                    }
                    //i = 0;
                    break;
                case 3:
                    labelLearningWord.Location = new Point((int)(0.5 * (this.Width - labelLearningWord.Width)), (int)(0.1 * this.Height));
                    buttonSubmit.Location = new Point((int)(0.25 * (this.Width - 2 * buttonSubmit.Width)), (int)(0.7 * this.Height));
                    buttonHint.Location = new Point((int)(0.75 * (this.Width - 2 * buttonHint.Width) + buttonHint.Width), (int)(0.7 * this.Height));
                    textBoxInput.Location = new Point((int)(0.5 * (this.Width - textBoxInput.Width)), (int)(0.5 * this.Height));
                    labelHint.Location = new Point(labelLearningWord.Location.X, textBoxInput.Location.Y + textBoxInput.Height + 50);
                    buttonNext.Location = buttonSubmit.Location;

                    break;
                default:
                    break;
            }
        }
    }
}


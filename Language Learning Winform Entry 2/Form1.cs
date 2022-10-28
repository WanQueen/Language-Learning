using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Language_Learning_Winform;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Language_Learning_Winform_Entry_2
{
    public partial class Form1 : Form
    {
        enum PAGE_NAME
        {
            NONE = 0,
            START_PAGE = 1,
            MODE_CHOSEN_PAGE = 2,
            LEARNING_PAGE = 3
        };
        //
        // MOST IMPORTANT VARIABLES
        //
        private string fileName = "";
        private int sheetNum = 0;
        private string[] sheetName;
        private string chosenSheetName = "";
        private DataTable dt = new DataTable();
        //
        // variables used for every page
        //
        private int page = (int)PAGE_NAME.NONE;
        private string wordNowLearning = "";
        private string wordCorrect = "";
        private string wordInput = "";
        private string wordHint = "";
        //
        // variables needed in start page
        //
        private Label labelQuest = new Label();
        private Button buttonChooseFile = new Button();
        //
        // variables needed in mode-chosen page
        //
        private Label labelQuest2 = new Label();
        private Panel panelModeChoose = new Panel();
        private List<Button> buttonModeChoose = new List<Button>();
        //
        // variables needed in learning page
        //
        private Label labelLearningWord = new Label();
        private Button buttonHint = new Button();
        private Button buttonSubmit = new Button();
        private Button buttonNext = new Button();
        private TextBox textBoxInput = new TextBox();
        private Label labelHint = new Label();
        //
        // variables for random word-choosing
        //
        private List<int> all = new List<int>();
        private Random rand = new Random();
        private int randNow = 0;


        //
        // function to change excel file to Datatable
        //
        public DataTable ExcelToDatatable(string fileName, string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            FileStream fs;
            IWorkbook workbook = null;
            int cellCount = 0;
            int rowCount = 0;

            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.IndexOf(".xls") > 0)
                {
                    workbook = new HSSFWorkbook(fs);
                }

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    cellCount = firstRow.LastCellNum; 
                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            firstRow.GetCell(i).SetCellType(CellType.String);
                            DataColumn column = new DataColumn(firstRow.GetCell(i).StringCellValue);
                            if (!data.Columns.Contains(firstRow.GetCell(i).StringCellValue))
                                data.Columns.Add(column);
                            else
                            {
                                column.ColumnName = column.ColumnName + i.ToString();
                                data.Columns.Add(column);
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }
                    
                    rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null)
                        {
                            continue;
                        }
                        
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) 
                            {
                                dataRow[j] = row.GetCell(j).ToString();
                            }
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        //
        // function to get excel file's sheet names
        //
        private void GetExcelSheetName()
        {
            IWorkbook workbook = null;
            FileStream fs;

            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            }
            catch (Exception ex)
            {
                this.Hide();
                MessageBox.Show(ex.Message);
                this.Close();
                Environment.Exit(Environment.ExitCode);
            }

            fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);

            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook(fs);
            }

            sheetNum = workbook.NumberOfSheets;
            sheetName = new string[sheetNum];
            for (int i = 0; i < sheetNum; i++)
                sheetName[i] = workbook.GetSheetName(i);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(821, 483);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Language Learning";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //
            // labelQuest
            //
            labelQuest.Width = 800;
            labelQuest.Height = 100;
            labelQuest.Location = new Point((int)(0.5 * (this.Width - labelQuest.Width)), (int)(0.1 * this.Height));
            labelQuest.TextAlign = ContentAlignment.MiddleCenter;
            labelQuest.Font = new Font("Arial", 20, FontStyle.Bold);
            labelQuest.Text = "Please choose the xls/xlsx file you wanna open";
            this.Controls.Add(labelQuest);
            //
            // buttonChooseFile
            //
            buttonChooseFile.Width = 200;
            buttonChooseFile.Height = 60;
            buttonChooseFile.Location = new Point((int)(0.5 * (this.Width - buttonChooseFile.Width)), (int)(0.7 * this.Height));
            buttonChooseFile.TextAlign = ContentAlignment.MiddleCenter;
            buttonChooseFile.Font = new Font("Arial", 10);
            buttonChooseFile.Text = "Click Here";
            buttonChooseFile.Click += ButtonChooseFile_Click;
            this.Controls.Add(buttonChooseFile);
        }

        private void ButtonChooseFile_Click(object sender, EventArgs e)
        {
            //
            // dialog
            //
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "Please choose the file";
            dialog.Filter = "Microsoft Excel file (*.xls, *.xlsx) | *.xls;*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                fileName = dialog.FileName;
                this.Controls.Remove(labelQuest);
                this.Controls.Remove(buttonChooseFile);
                //
                // enter the interface of choosing mode
                //
                ModeChoose();
            }
        }

        private void ModeChoose()
        {
            page = (int)PAGE_NAME.MODE_CHOSEN_PAGE;

            GetExcelSheetName();
            //
            // labelQuest2
            //
            labelQuest2.Width = 800;
            labelQuest2.Height = 100;
            labelQuest2.Location = new Point((int)(0.5 * (this.Width - labelQuest2.Width)), (int)(0.1 * this.Height));
            labelQuest2.TextAlign = ContentAlignment.MiddleCenter;
            labelQuest2.Font = new Font("Arial", 20, FontStyle.Bold);
            labelQuest2.Text = "Please choose the sheet you wanna learn";
            this.Controls.Add(labelQuest2);
            //
            // panelModeChoose
            //
            panelModeChoose.Width = this.Width;
            panelModeChoose.Height = (int)(this.Height - 150);
            panelModeChoose.Location = new Point(0, (int)(0.15 * this.Height));
            panelModeChoose.AutoScroll = true;
            this.Controls.Add(panelModeChoose);
            //
            // buttonModeChoose
            //
            int allButtonLocationX = 50;
            int allButtonLocationY = labelQuest2.Height;
            for (int i = 0; i < sheetNum; i++)
            {
                Button btn = new Button();

                //btn.Width = 100;
                //btn.Height = 50;
                btn.AutoSize = true;
                //btn.Location = new Point((int)(0.5 * (this.Width - btn.Width)), labelQuest2.Height + (int)(0.1 * (i + 1) * this.Height));
                btn.TextAlign = ContentAlignment.MiddleCenter;
                btn.Font = new Font("Arial", 20);
                btn.Text = sheetName[i];
                btn.Visible = false;
                panelModeChoose.Controls.Add(btn);
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
                btn.Visible = true;
                btn.Click += ButtonModeChoose_Click;
                buttonModeChoose.Add(btn);
            }
        }

        private void ButtonModeChoose_Click(object sender, EventArgs e)
        {
            chosenSheetName = ((Button)sender).Text;
            dt = ExcelToDatatable(fileName, chosenSheetName, true);
            Learning();
        }

        private void Learning()
        {
            page = (int)PAGE_NAME.LEARNING_PAGE;

            this.Controls.Remove(labelQuest2);
            foreach (Button btn in buttonModeChoose)
            {
                panelModeChoose.Controls.Remove(btn);
            }
            this.Controls.Remove(panelModeChoose);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                all.Add(i);
            }

            RandomWordChoose();
            wordNowLearning = dt.Rows[randNow][1].ToString();
            wordCorrect = dt.Rows[randNow][0].ToString();
            wordHint = dt.Rows[randNow][2].ToString();
            //
            // labelLearningWord
            //
            labelLearningWord.Width = 800;
            labelLearningWord.Height = 100;
            labelLearningWord.Location = new Point((int)(0.5 * (this.Width - labelLearningWord.Width)), (int)(0.1 * this.Height));
            labelLearningWord.TextAlign = ContentAlignment.MiddleCenter;
            labelLearningWord.Font = new Font("Arial", 40, FontStyle.Bold);
            labelLearningWord.Text = wordNowLearning;
            this.Controls.Add(labelLearningWord);
            //
            // buttonSubmit
            //
            buttonSubmit.Width = 200;
            buttonSubmit.Height = 50;
            buttonSubmit.Location = new Point((int)(0.25 * (this.Width - 2 * buttonSubmit.Width)), (int)(0.7 * this.Height));
            buttonSubmit.TextAlign = ContentAlignment.MiddleCenter;
            buttonSubmit.Font = new Font("Arial", 20);
            buttonSubmit.Text = "Submit";
            buttonSubmit.Click += buttonSubmit_Click;
            this.Controls.Add(buttonSubmit);
            //
            // buttonHint
            //
            buttonHint.Width = 200;
            buttonHint.Height = 50;
            buttonHint.Location = new Point((int)(0.75 * (this.Width - 2 * buttonHint.Width) + buttonHint.Width), (int)(0.7 * this.Height));
            buttonHint.TextAlign = ContentAlignment.MiddleCenter;
            buttonHint.Font = new Font("Arial", 20);
            buttonHint.Text = "Hint";
            buttonHint.Click += buttonHint_Click;
            this.Controls.Add(buttonHint);
            //
            // textBoxInput 
            //
            textBoxInput.Width = (int)(0.7 * this.Width);
            textBoxInput.Height = 100;
            textBoxInput.Location = new Point((int)(0.5 * (this.Width - textBoxInput.Width)), (int)(0.5 * this.Height));
            textBoxInput.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            textBoxInput.Font = new Font("Arial", 20);
            textBoxInput.TextChanged += textBoxInput_TextChanged;
            this.Controls.Add(textBoxInput);
            //
            // labelHint
            //
            labelHint.Width = labelLearningWord.Width;
            labelHint.Height = labelLearningWord.Height;
            labelHint.Location = new Point(labelLearningWord.Location.X, textBoxInput.Location.Y + textBoxInput.Height);
            labelHint.TextAlign = ContentAlignment.MiddleCenter;
            labelHint.Font = new Font("Arial", 10);
            labelHint.Text = wordHint;
            labelHint.Visible = false;
            this.Controls.Add(labelHint);
            //
            // buttonNext
            //
            buttonNext.Width = (int)(0.5 * (this.Width - 2 * buttonHint.Width)) + buttonHint.Width * 2;
            buttonNext.Height = 50;
            buttonNext.Location = buttonSubmit.Location;
            buttonNext.TextAlign = ContentAlignment.MiddleCenter;
            buttonNext.Font = new Font("Arial", 20);
            buttonNext.Text = "Next";
            buttonNext.Click += buttonNext_Click;
            buttonNext.Visible = false;
            this.Controls.Add(buttonNext);
        }

        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            if (wordInput == wordCorrect)
            {
                labelHint.Visible = false;

                this.BackColor = Color.FromArgb(46, 139, 87);
                textBoxInput.BackColor = Color.FromArgb(127, 255, 170);

                buttonSubmit.Visible = false;
                buttonHint.Visible = false;
                buttonNext.Visible = true;
            }
            else
            {
                this.BackColor = Color.FromArgb(255, 69, 0);
                textBoxInput.BackColor = Color.FromArgb(255, 99, 71);
            }
        }

        private void buttonHint_Click(object sender, EventArgs e)
        {
            labelHint.Visible = true;
        }

        private void textBoxInput_TextChanged(object sender, EventArgs e)
        {
            wordInput = ((TextBox)sender).Text;
        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            RandomWordChoose();
            wordNowLearning = dt.Rows[randNow][1].ToString();
            wordCorrect = dt.Rows[randNow][0].ToString();
            wordHint = dt.Rows[randNow][2].ToString();

            labelLearningWord.Text = wordNowLearning;
            labelHint.Text = wordHint;

            this.BackColor = Color.White;
            textBoxInput.BackColor = Color.White;
            textBoxInput.Text = "";

            if (all.Count > 0)
            {
                buttonNext.Visible = false;
                buttonSubmit.Visible = true;
                buttonHint.Visible = true;
                labelHint.Visible = false;
            }
            else
            {
                buttonNext.Visible = false;
                buttonSubmit.Visible = false;
                buttonHint.Visible = false;
                labelHint.Visible = false;
                textBoxInput.Visible = false;
                labelLearningWord.Text = "You finished all the words!";
            }
        }

        private void RandomWordChoose()
        {
            int index = new int();
            if (all.Count > 0)
                index = rand.Next(0, all.Count - 1);
            else
                return;
            randNow = all[index];
            all.RemoveAt(index);
        }
    }
}

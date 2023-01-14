using ExcelLibrary.BinaryFileFormat;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace _1234
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.Filter="Excel Files|*.xlsx";

                openFileDialog1.ShowDialog();

                MessageBox.Show("Выберите папку в которую сохранить");

                folderBrowserDialog1.ShowDialog();

                string pathToSave = folderBrowserDialog1.SelectedPath;

                string filename = openFileDialog1.FileName;

                if (filename != "openFileDialog1" && pathToSave!="")
                {
                                    

                FileInfo fileInfo = new FileInfo(filename);

                ExcelPackage package = new ExcelPackage(fileInfo);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                using (ExcelPackage excel = new ExcelPackage())
                {
                    excel.Workbook.Worksheets.Add("Worksheet1");
                    FileInfo excelFile = new FileInfo(pathToSave + "\\forPlatform.xlsx");
                    excel.SaveAs(excelFile);
                }

                using (ExcelPackage excel = new ExcelPackage())
                {
                    excel.Workbook.Worksheets.Add("Worksheet1");
                    FileInfo excelFile2 = new FileInfo(pathToSave + "\\forOperator.xlsx");
                    excel.SaveAs(excelFile2);
                }

                ExcelPackage package2 = new ExcelPackage(pathToSave + "\\forPlatform.xlsx");
                ExcelWorksheet worksheet2 = package2.Workbook.Worksheets.FirstOrDefault();

                ExcelPackage package3 = new ExcelPackage(pathToSave + "\\forOperator.xlsx");
                ExcelWorksheet worksheet3 = package3.Workbook.Worksheets.FirstOrDefault();


                // get number of rows in the sheet
                int rows = worksheet.Dimension.Rows; // 10

                int sot = 2;
                int dom = 2;
                int musor = 2;
                // loop through the worksheet rows
                for (int i = 2; i <= rows; i++)
                {
                    if (worksheet.Cells[i, 5].Value != null)
                    {
                        string newNubmer = worksheet.Cells[i, 5].Value.ToString();//access specific cell and modify its value

                        newNubmer = Regex.Replace(newNubmer, "[^-.0-9]", "");

                        string newbumber = "";
                        newbumber = newNubmer.Replace("-", "").Replace("+", "").Replace("(", "").Replace(")", "").Replace(" ", "");
                        if (newbumber.Length == 11)
                        {
                            newbumber = newbumber.Replace(newNubmer.First().ToString(), "8");
                            worksheet2.Cells[sot, 1].Value = (newbumber);
                            package2.Save();
                            sot++;
                        }
                        else if (newbumber.Length == 10)
                        {
                            newbumber = "8" + newbumber;
                            worksheet2.Cells[sot, 1].Value = (newbumber);
                            package2.Save();
                            sot++;

                        }
                        else if (newbumber.Length == 7)
                        {
                            worksheet3.Cells[dom, 1].Value = (newbumber);
                            package3.Save();
                            dom++;

                        }

                        else
                        {

                            worksheet3.Cells[musor, 2].Value = worksheet.Cells[i, 5].Value.ToString();
                            package3.Save();
                            musor++;


                        }

                    }


                }

                worksheet2.Cells[1, 1].Value = ("Номера");
                worksheet3.Cells[1, 1].Value = ("Домашние телефоны");
                worksheet3.Cells[1, 2].Value = ("Сборные номера");

                package2.Save();
                package3.Save();



                MessageBox.Show("Все готово!");

            }
                else
                {
                    MessageBox.Show("Файл не выбран");
                }
            }
            catch (Exception ex)

            {

                MessageBox.Show("Данные некорректны");

            }


        }
    }
}

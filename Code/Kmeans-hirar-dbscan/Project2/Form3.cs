using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CenterSpace.NMath.Matrix;
using CenterSpace.NMath.Core;
using Extreme.Mathematics;
using Extreme.Mathematics.LinearAlgebra.IO;
using Extreme.Statistics;
using Extreme.Statistics.Multivariate;
using Microsoft.Office.Interop.Excel;
namespace Project2
{
    public partial class Form3 : Form
    {
        String[] data = new String[100];
        double[,] attr = new double[1000, 17];
        int[,] gt = new int[1000, 3];
        string[,] g_cl_temp = new string[1000, 2];
        string[,] g_cl = new string[1000, 2];
        double[,] g_mean = new double[5, 17];
        int rows = 0;
        int cols = 0;
        int cluster_count = 0;
        private BindingSource bindingSource1 = new BindingSource();
        private BindingSource bindingSource2 = new BindingSource();
        Dictionary<string, double> dm1 = new Dictionary<string, double>();
        public Form3()
        {
            InitializeComponent();
        }

        private void Hier(string[,] cluster, int itr, double max_val, string max_key)
        {
            string[,] cl = new string[rows - itr, rows - itr];
            string[] words = max_key.Split('-');
            int col = 1;
            
            for (int k = 0; k < words.Length; k++)
            {
                if (k == 0)
                {
                    cl[0, 1] = words[k];
                }
                else
                {
                    cl[0, 1] = cl[0, 1] + "," + words[k];
                }
            }
            col = 2;
            for (int j = 1; j <= rows - itr; j++)
            {
                int flag_k = 0;
                for (int k = 0; k < words.Length; k++)
                {
                    if (cluster[0, j] == words[k])
                    {
                        flag_k++;
                    }


                }
                if (flag_k != 0)
                {
                    continue;
                }
                cl[0, col] = cluster[0, j];
                col++;

            }
            double t_max = 99;
            string t_key = null;
            for (int i = 1; i < rows - itr; i++)
            {
                string[] words_i = cl[0, i].Split(',');
                for (int k = 0; k < words_i.Length; k++)
                {
                    for (int j = i + 1; j < rows - itr; j++)
                    {
                        string[] words_j = cl[0, j].Split(',');
                        for (int l = 0; l < words_j.Length; l++)
                        {
                            double val = 0;
                            if (dm1.TryGetValue(words_i[k] + "-" + words_j[l], out val))
                            {
                                if (val < t_max)
                                {
                                    t_max = val;
                                    t_key = cl[0, i] + "-" + cl[0, j];
                                }
                            }
                        }
                    }
                }
            }



            itr++;
            if (cluster_count == rows-itr)
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 1; j <= cluster_count; j++)
                    {
                        string[] words_new = cl[0, j].Split(',');
                        for (int k = 0; k < words_new.Length; k++)
                        {
                            g_cl[Convert.ToInt32(words_new[k])-1, 0] = words_new[k];
                            g_cl[Convert.ToInt32(words_new[k]) -1, 1] = j.ToString();
                        }
                    }
                    

                }
                return;
            }
            Hier(cl, itr, t_max, t_key);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int itr = 0;
            string[,] cluster = new string[1000, 1000];
            Microsoft.Office.Interop.Excel.Application IExcel = new Microsoft.Office.Interop.Excel.Application();
            //IExcel.Visible = true;
            string filepath = null;


            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                filepath = openFileDialog1.FileName;
            }

            cluster_count = Convert.ToInt32(textBox1.Text);
            //string fileName = "C:\\Users\\Avijeet\\Desktop\\Data Mining 601\\Project 2\\cho.xlsx";
            string fileName = filepath;
            //open the workbook
            Workbook workbook = IExcel.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            //select the first sheet        
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            //find the used range in worksheet
            Microsoft.Office.Interop.Excel.Range excelRange = worksheet.UsedRange;

            //get an object array of all of the cells in the worksheet (their values)
            object[,] valueArray = (object[,])excelRange.get_Value(
                        XlRangeValueDataType.xlRangeValueDefault);
            rows = worksheet.UsedRange.Rows.Count;
            cols = worksheet.UsedRange.Columns.Count;
            for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
            {
                for (int col = 1; col < 3; ++col)
                {
                    //access each cell
                    gt[row, col] = Convert.ToInt32(valueArray[row, col]);
                }
            }

            //access the cells
            for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
            {
                for (int col = 3; col <= worksheet.UsedRange.Columns.Count; ++col)
                {
                    //access each cell
                    attr[row, col - 2] = Convert.ToDouble(valueArray[row, col]);
                }
            }

            //clean up stuffs
            workbook.Close(false, Type.Missing, Type.Missing);

            IExcel.Quit();
            double max = 99;
            string max_key = null;
            //Create Distance Matrix
            for (int i = 1; i < rows + 1; i++)
            {
                cluster[i, 0] = i.ToString();
                for (int j = 1; j < rows + 1; j++)
                {
                    cluster[0, j] = j.ToString();
                    double sum = 0;
                    string key = i + "-" + j;
                    if (i == j)
                    {
                        dm1.Add(key, 1);
                        cluster[i, j] = 1.ToString();
                    }
                    else
                    {
                        for (int k = 1; k < cols-1; k++)
                        {
                            sum = sum + Math.Pow(attr[i, k] - attr[j, k], 2);

                        }
                        //double sim_val = 1 / (1 + Math.Sqrt(sum));
                        double sim_val = Math.Sqrt(sum);
                        dm1.Add(key, sim_val);
                        cluster[i, j] = sim_val.ToString();
                        if (sim_val < max)
                        {
                            max = sim_val;
                            max_key = key;
                        }
                    }

                }

            }

            Hier(cluster, itr, max, max_key);
            int[,] gd_matx = new int[rows + 1, rows + 1];
            int[,] cl_matx = new int[rows + 1, rows + 1];
            double m11 = 0, m00 = 0, m01 = 0, m10 = 0;
            for (int i = 1; i < rows + 1; i++)
            {
                for (int j = 1; j < rows + 1; j++)
                {
                    if (gt[i, 2] == gt[j, 2])
                    {
                        gd_matx[i, j] = 1;
                    }
                    else
                    {
                        gd_matx[i, j] = 0;
                    }
                    if (g_cl[i - 1, 1] == g_cl[j - 1, 1])
                    {
                        cl_matx[i, j] = 1;
                        if (gd_matx[i, j] == 1)
                        {
                            m11++;
                        }
                        else if (gd_matx[i, j] == 0)
                        {
                            m10++;
                        }
                    }
                    else
                    {
                        cl_matx[i, j] = 0;
                        if (gd_matx[i, j] == 1)
                        {
                            m01++;
                        }
                        else if (gd_matx[i, j] == 0)
                        {
                            m00++;
                        }
                    }

                }
            }

            double rand;
            rand = Convert.ToDouble((m11 + m00) / (m11 + m00 + m01 + m10));
            double jaccof;
            jaccof = Convert.ToDouble((m11) / (m11 + m01 + m10));

            label5.Text = label5.Text + jaccof.ToString();
            label6.Text = label6.Text + rand.ToString();

            int[,] res_arr = new int[cluster_count, 2];
            for (int i = 0; i < cluster_count; i++)
            {
                res_arr[i, 1] = 0;
                for (int j = 0; j < g_cl.GetLength(0) - 1; j++)
                {
                    if (g_cl[j, 1] == (i+1).ToString())
                    {
                        res_arr[i, 0] = i + 1;
                        res_arr[i, 1] = res_arr[i, 1] + 1;
                    }
                }
            }
            dataGridView1.ColumnCount = 2;

            for (int i = 0; i < res_arr.GetLength(0); i++)
            {
                dataGridView1.Rows.Add(new object[] { res_arr[i, 0], res_arr[i, 1] });
            }
            //Cluster assignment array

            dataGridView2.ColumnCount = cols;

            for (int j = 1; j < attr.GetLength(0); j++) 
            {
                    if (j.ToString() == g_cl[j-1,0])
                    {
                    if (cols == 18)
                    {
                        dataGridView2.Rows.Add(new object[] { j, g_cl[j - 1, 1], attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5], attr[j, 6], attr[j, 7], attr[j, 8], attr[j, 9], attr[j, 10], attr[j, 11], attr[j, 12], attr[j, 13], attr[j, 14], attr[j, 15], attr[j, 16] });
                    }
                    else if (cols == 14)
                    {
                        dataGridView2.Rows.Add(new object[] { j, g_cl[j - 1, 1], attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5], attr[j, 6], attr[j, 7], attr[j, 8], attr[j, 9], attr[j, 10], attr[j, 11], attr[j, 12]});
                    }
                    else if (cols == 6)
                    {
                        dataGridView2.Rows.Add(new object[] { j, g_cl[j - 1, 1], attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4]});
                    }
                    else if (cols == 7)
                    {
                        dataGridView2.Rows.Add(new object[] { j, g_cl[j - 1, 1], attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5] });
                    }
                }

            }



        }
    }
}

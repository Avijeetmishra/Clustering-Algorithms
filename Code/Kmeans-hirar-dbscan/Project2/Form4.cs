using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using CenterSpace.NMath.Matrix;
using CenterSpace.NMath.Core;
using Extreme.Mathematics;
using Extreme.Mathematics.LinearAlgebra.IO;
using Extreme.Statistics;
using Extreme.Statistics.Multivariate;
using System.Text.RegularExpressions;

namespace Project2
{
    public partial class Form4 : Form
    {
        String[] data = new String[100];
        double[,] attr = new double[1000, 17];
        int[,] gt = new int[1000, 4];
        string[,] g_cl = new string[1000, 2];
        double[,] g_mean = new double[30, 17];
        int cluster_count = 0;
        int rows = 0;
        int cols = 0;
        string[,] cluster = new string[1000, 2];
        string[,] noise = new string[1, 2];
        int noise_count = 0;
        

        private BindingSource bindingSource1 = new BindingSource();
        private BindingSource bindingSource2 = new BindingSource();
        Dictionary<string, double> dm1 = new Dictionary<string, double>();
        public Form4()
        {
            InitializeComponent();
        }

        private void DBSCAN(int[,] D, double eps, int MinPts)
        {
            int C = 0;
            //'int[,] Neighborpts = new int[rows, 2];
            for (int i = 1; i <= rows; i++)
            {
                
                if (D[i, 3] == 0)
                {
                    D[i, 3] = 1;
                    int[] Neighborpts = regionQuery(D[i,1], D, eps);
                    if (Neighborpts.GetLength(0) < MinPts)
                    {
                        D[i, 3] = -1;
                        noise_count++;
                        noise[0, 1] = noise[0, 1] + "," + D[i,1].ToString();
                    }
                    else
                    {
                        C++;
                        expandCluster(D, D[i, 1], Neighborpts, C, eps, MinPts);
                    }

                }
               
            }
            cluster_count = C;
        }

        private void expandCluster(int[,] D, int p, int[] Neighborpts, int C, double eps, int MinPts)
        {
            cluster[C, 0] = C.ToString();
            cluster[C, 1] = "," + p.ToString() + ",";
            int col=1;
            for (int i = 0; i < Neighborpts.Length; i++)
            {
                
                int p1 = Neighborpts[i];
                for (int j = 0; j < D.GetLength(0); j++)
                {
                    if (D[j, 1] == p1)
                    {
                        if (D[j, 3] == 0)
                        {
                            D[j, 3] = 1;
                            int[] Neighborpts1 = regionQuery(p1, D, eps);
                            if (Neighborpts1.Length >=MinPts)
                            {
                                for(int x = 0; x < Neighborpts1.Length; x++)
                                {
                                    int flag = 0;
                                    for(int y = 0; y < Neighborpts.Length; y++)
                                    {
                                        if (Neighborpts1[x] == Neighborpts[y])
                                        {
                                            flag= 1;
                                        }
                                    }
                                    if (flag ==0)
                                    {
                                        Array.Resize(ref Neighborpts, Neighborpts.Length+1);
                                        Neighborpts[Neighborpts.Length-1] = Neighborpts1[x];
                                    }
                                }
                            }

                            
                        }
                    }

                    
                }
                int flag_p = 0;
                for (int k = 1; k <=C;k++)
                {
                    
                        if(cluster[k, 1].Contains("," + p1.ToString() +","))
                        {
                            flag_p = 1;
                        }

                }
                if (flag_p == 0)
                {
                    cluster[C, 1] = cluster[C, 1] + "," + p1.ToString() + ",";
                    col++;
                }
            }
        }
        //return all points within P's eps-neighborhood (including P)
        private int[] regionQuery(int p, int[,] D, double eps)
        {   int[] nbr = new int[rows];
            
            int count = 0;
            for (int i = 0; i <= rows; i++)
            {   string key = null;
                if (p < D[i, 1])
                {
                    key = p + "-" + D[i, 1];
                }
                else
                {
                    key = D[i, 1] + "-" + p;
                }
                double val = 0;
                if (dm1.TryGetValue(key, out val))
                {
                    if (val >= eps)
                    {
                        nbr[count] = D[i, 1];
                        count++;
                    }

                }
            }
            Array.Resize(ref nbr, count);
            return nbr;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filepath = null;
            noise[0, 0] = "Noise";

            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                filepath = openFileDialog1.FileName;
            }
            
            Microsoft.Office.Interop.Excel.Application IExcel = new Microsoft.Office.Interop.Excel.Application();
            ///IExcel.Visible = true;
           


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
            rows = worksheet.UsedRange.Rows.Count;
            cols = worksheet.UsedRange.Columns.Count;
            //clean up stuffs
            workbook.Close(false, Type.Missing, Type.Missing);
            IExcel.Quit();
            double eps = Convert.ToDouble(textBox1.Text);
            int minpts = Convert.ToInt32(textBox2.Text);

            //Create Distance Matrix
            for (int i = 1; i < rows+1; i++)
            {
                for (int j = 1; j < rows+1; j++)
                {
                    double sum = 0;
                    string key = i + "-" + j;
                    if (i == j)
                    {
                        dm1.Add(key, 1);
                    }
                    else
                    {
                        for (int k = 1; k < 17; k++)
                        {
                            sum = sum + Math.Pow(attr[i, k] - attr[j, k], 2);

                        }
                        double sim_val = 1 / (1 + Math.Sqrt(sum));
                        dm1.Add(key, sim_val);
                    }

                }

            }
            DBSCAN(gt, eps , minpts);

            for (int i = 1; i <= cluster_count; i++)
            {
                cluster[i, 1] = cluster[i, 1].Substring(1, cluster[i, 1].Length-2);
            }

            for (int i = 0; i < rows; i++)
            {
                for (int j = 1; j <= cluster_count; j++)
                {
                    string[] words_new = Regex.Split(cluster[j, 1], ",,");
                    for (int k = 0; k < words_new.Length; k++)
                    {
                        g_cl[Convert.ToInt32(words_new[k]) - 1, 0] = words_new[k];
                        g_cl[Convert.ToInt32(words_new[k]) - 1, 1] = j.ToString();
                    }
                }


            }

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
                    if (g_cl[j, 1] == (i + 1).ToString())
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
            dataGridView1.Rows.Add(new object[] { -1, noise_count });

            dataGridView2.ColumnCount = cols;

            for (int j = 1; j < attr.GetLength(0); j++)
            {
                if (j.ToString() == g_cl[j - 1, 0])
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
                        dataGridView2.Rows.Add(new object[] { j, g_cl[j - 1, 1], attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4] });
                    }
                    else if (cols == 7)
                    {
                        dataGridView2.Rows.Add(new object[] { j, g_cl[j - 1, 1], attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5]});
                    }
                }

            }
            string[,] g_noise = new string[noise_count, 2];

            if (noise[0, 1] != null)
            {
                string[] words_new1 = Regex.Split(noise[0, 1], ",");
                for (int i = 1; i < words_new1.Length; i++)
                {

                    g_noise[i - 1, 0] = words_new1[i];
                    g_noise[i - 1, 1] = (-1).ToString();

                }

                dataGridView3.ColumnCount = cols;

                for (int j = 1; j <= rows; j++)
                {
                    for (int i = 0; i < noise_count; i++)
                    {
                        if (j.ToString() == g_noise[i, 0])
                        {
                            if (cols == 18)
                            {
                                dataGridView3.Rows.Add(new object[] { j, -1, attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5], attr[j, 6], attr[j, 7], attr[j, 8], attr[j, 9], attr[j, 10], attr[j, 11], attr[j, 12], attr[j, 13], attr[j, 14], attr[j, 15], attr[j, 16] });
                            }
                            else if (cols == 14)
                            {
                                dataGridView3.Rows.Add(new object[] { j, -1, attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5], attr[j, 6], attr[j, 7], attr[j, 8], attr[j, 9], attr[j, 10], attr[j, 11], attr[j, 12]});
                            }
                            else if (cols == 6)
                            {
                                dataGridView3.Rows.Add(new object[] { j, -1, attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4] });
                            }
                            else if (cols == 7)
                            {
                                dataGridView3.Rows.Add(new object[] { j, -1, attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5] });
                            }
                        }
                    }

                }
            } 

        }
    }
}

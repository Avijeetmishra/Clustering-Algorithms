using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using Microsoft.Office.Interop.Excel;
using CenterSpace.NMath.Matrix;
using CenterSpace.NMath.Core;
using Extreme.Mathematics;
using Extreme.Mathematics.LinearAlgebra.IO;
using Extreme.Statistics;
using Extreme.Statistics.Multivariate;

namespace Project2
{
    public partial class Form2 : Form
    {
        String[] data = new String[100];
        double[,] attr = new double[1000, 17];
        int[,] gt = new int[1000, 3];
        int[,] g_cl = new int[1000, 2];
        double[,] g_mean = new double[30, 17];
        int cluster_count = 0;
        int g_itr = 0;
        int rows = 0;
        int cols = 0;
        int input_itr = 0;
        private BindingSource bindingSource1 = new BindingSource();
        private BindingSource bindingSource2 = new BindingSource();
        Dictionary<string, double> dm1 = new Dictionary<string, double>();
        public Form2()
        {
            InitializeComponent();

        }
        private void kmeans(double[,] cluster, int itr)
        {
            int[,] cl = new int[rows + 1, 2];
            double[,] mean = new double[cluster_count, rows + 1];
            Dictionary<string, double> dm = new Dictionary<string, double>();


            for (int i = 1; i < rows + 1; i++)
            {
                for (int j = 0; j < cluster_count; j++)
                {
                    double sum = 0;
                    double val = cluster[j, 0];
                    for (int k = 1; k < cols-1; k++)
                    {
                        if (itr == 0)
                        {
                            sum = sum + Math.Pow(attr[i, k] - attr[Convert.ToInt32(val), k], 2);
                        }
                        else
                        {
                            sum = sum + Math.Pow(attr[i, k] - cluster[j, k], 2);
                        }
                    }

                    dm.Add(i + "," + Convert.ToInt32(val), 1 / (1 + Math.Sqrt(sum)));

                }

            }

            int count = 0;
            double min = 0;
            string key = null;
            int index = 0;
            foreach (KeyValuePair<string, double> kp in dm)
            {
                count++;
                if (kp.Value > min)
                {
                    min = kp.Value;
                    key = kp.Key;
                }
                if (count == cluster_count)
                {
                    string[] words = key.Split(',');
                    cl[index, 0] = Convert.ToInt32(words[0]);
                    cl[index, 1] = Convert.ToInt32(words[1]);
                    index++;
                    min = 0;
                    count = 0;
                    key = null;
                }

            }
            //Calculate mean for centroid

            for (int j = 0; j < cluster_count; j++)
            {
                count = 0;
                for (int i = 0; i < rows; i++)
                {
                    if (cl[i, 1] == cluster[j, 0])
                    {
                        int id = cl[i, 0];
                        mean[j, 0] = j;
                        for (int k = 1; k < cols-1; k++)
                        {
                            mean[j, k] = mean[j, k] + attr[id, k];
                            if (count == 0)
                            {
                                if (itr == 0)
                                {
                                    mean[j, k] = mean[j, k] + attr[Convert.ToInt32(cluster[j, 0]), k];
                                }
                                else
                                {
                                    mean[j, k] = mean[j, k] + cluster[j, k];
                                }
                            }

                        }
                        count++;
                    }
                }

                for (int k = 1; k < cols - 1; k++)
                {

                    mean[j, k] = mean[j, k] / (count + 1);
                }

            }
            count = 0;
            int count1 = 0;
            for (int j = 0; j < cluster_count; j++)
            {
                for (int k = 1; k < cols - 1; k++)
                {
                    if (mean[j, k] == cluster[j, k])
                    {
                        count1++;
                    }
                }
                if (count1 == cols - 2)
                {
                    count++;
                    count1 = 0;
                }

            }
            if (count == cluster_count)
            {
                g_mean = mean;
                g_cl = cl;
                g_itr = itr;
                return;
            }
            itr++;
            if (input_itr != 0)
            {
                if (itr == input_itr)
                {
                    g_mean = mean;
                    g_cl = cl;
                    g_itr = itr;
                    return;
                }
            }
            kmeans(mean, itr);

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filepath = null;
 
           
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                filepath = openFileDialog1.FileName;
            }
            int itr = 0;
            
            Microsoft.Office.Interop.Excel.Application IExcel = new Microsoft.Office.Interop.Excel.Application();
            ///IExcel.Visible = true;
            cluster_count = Convert.ToInt32(textBox1.Text);
            if (cluster_count == 0)
            {
                Random r = new Random();
                cluster_count = r.Next(3, 10);
            }

            double[,] cluster = new double[cluster_count, 17];
            if (textBox2.Text == "")
            {
                Random r = new Random();

                for (int i = 0; i < cluster_count; i++)
                {
                    cluster[i, 0] = r.Next(1, 386);
                }
            }
            else
            {

                string[] words = textBox2.Text.Split(',');

                for (int i = 0; i < cluster_count; i++)
                {
                    cluster[i, 0] = Convert.ToDouble(words[i]);
                }

            }

            if (textBox3.Text != "")
            {
                input_itr = Convert.ToInt32(textBox3.Text);
            }

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
            kmeans(cluster, itr);
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
            for (int i = 0; i < cluster.GetLength(0); i++)
            {
                res_arr[i, 1] = 0;
                for (int j = 0; j < g_cl.GetLength(0)-1; j++)
                {
                    if (g_cl[j, 1]==i)
                    {
                        res_arr[i, 0] = i+1;
                        res_arr[i, 1] = res_arr[i, 1] + 1;
                    } 
                }
            }
            dataGridView1.ColumnCount = 2;
            
            for (int i = 0; i < res_arr.GetLength(0); i++)
            {
                dataGridView1.Rows.Add(new object[] { res_arr[i, 0], res_arr[i,1] });
            }
            //Cluster assignment array

            dataGridView2.ColumnCount = cols;

            for (int i = 0; i < g_cl.GetLength(0); i++)
            {
                for (int j = 1; j < attr.GetLength(0); j++)
                {
                    if (g_cl[i, 0] == j)
                    {
                        if (cols == 18)
                        {
                            dataGridView2.Rows.Add(new object[] { g_cl[i, 0], g_cl[i, 1] + 1, attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5], attr[j, 6], attr[j, 7], attr[j, 8], attr[j, 9], attr[j, 10], attr[j, 11], attr[j, 12], attr[j, 13], attr[j, 14], attr[j, 15], attr[j, 16] });
                        }
                        else if (cols == 14)
                        {
                            dataGridView2.Rows.Add(new object[] { g_cl[i, 0], g_cl[i, 1] + 1, attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j, 5], attr[j, 6], attr[j, 7], attr[j, 8], attr[j, 9], attr[j, 10], attr[j, 11], attr[j, 12] });
                        }
                        else if (cols == 6)
                        {
                            dataGridView2.Rows.Add(new object[] { g_cl[i, 0], g_cl[i, 1] + 1, attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4]});
                        }
                        else if (cols == 7)
                        {
                            dataGridView2.Rows.Add(new object[] { g_cl[i, 0], g_cl[i, 1] + 1, attr[j, 1], attr[j, 2], attr[j, 3], attr[j, 4], attr[j,5] });
                        }
                    } 
                }
                
            }


            //PCA
            //Matrix 
            /*var m = Matrix.Create(g_cl);
            //bindingSource1.DataSource = m;
            //dataGridView1.DataSource = bindingSource1;
            //bindingSource2.DataSource = g_cl;
            //dataGridView2.DataSource = bindingSource2;
            PrincipalComponentAnalysis pca = new PrincipalComponentAnalysis(m);
            pca.Compute();
            // We can get the contributions of each component:
            Console.WriteLine(" #    Eigenvalue Difference Contribution Contrib. %");
            for (int i = 0; i < 2; i++)
            {
                // We get the ith component from the model...
                PrincipalComponent component = pca.Components[i];
                // and write out its properties
                Console.WriteLine("{0,2}{1,12:F4}{2,11:F4}{2,14:F3}%{3,10:F3}%",
                    i, component.Eigenvalue, component.EigenvalueDifference,
                    100 * component.ProportionOfVariance,
                    100 * component.CumulativeProportionOfVariance);
            }

            // To get the proportions for all components, use the
            // properties of the PCA object:
            var proportions = pca.VarianceProportions;

            // To get the number of components that explain a given proportion
            // of the variation, use the GetVarianceThreshold method:
            int count = pca.GetVarianceThreshold(0.9);
            Console.WriteLine("Components needed to explain 90% of variation: {0}", count);
            Console.WriteLine();

            // The value property gives the components themselves:
            Console.WriteLine("Components:");
            Console.WriteLine("Var.      1       2       3       4       5");
            PrincipalComponentCollection pcs = pca.Components;
            for (int i = 0; i < pcs.Count; i++)
            {

                Console.WriteLine("{0,4}{1,8:F4}{2,8:F4}",
                    i, pcs[0].Value[i], pcs[1].Value[i]);
            }
            Console.WriteLine();

            // The scores are the coefficients of the observations expressed as a combination
            // of principal components.
            var scores = pca.ScoreMatrix;

            // To get the predicted observations based on a specified number of components,
            // use the GetPredictions method.
            var prediction = pca.GetPredictions(count);
            Console.WriteLine("Predictions using {0} components:", count);
            Console.WriteLine("   Pr. 1  Act. 1   Pr. 2  Act. 2   Pr. 3  Act. 3   Pr. 4  Act. 4", count);
           /* for (int i = 0; i < 10; i++)
                Console.WriteLine("{0,8:F4}{1,8:F4}{2,8:F4}{3,8:F4}{4,8:F4}{5,8:F4}{6,8:F4}{7,8:F4}",
                    prediction[i, 0], m[i, 0],
                    prediction[i, 1], m[i, 1]);*/


            
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest_1(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}

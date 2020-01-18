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
using System.IO;
using System.Windows.Forms.DataVisualization.Charting;

namespace OpTiCs
{
    public partial class Form1 : Form
    {
        static int MinPts;
        static double Eps;
        static double secondEps;
        static int RowNum;
        static int ColNum;
        int[,] A;
        class Object
        {
            public int Name;
            public double B;
            public double G;
            public double R;
            public bool Processed = false;
            public double reachability_distance = -1; //UNDEFINED
            public double core_distance = -1; //UNDEFINED
            public int ClusterId;
            public Object(double _B, double _G, double _R, int _Name)
            {
                B = _B;
                G = _G;
                R = _R;
                Name = _Name;
            }
        }
        class setOfObjects : List<Object>
        {
        }
        #region Read Exel
        static int[,] Read(string S)
        {
            string d = Directory.GetCurrentDirectory(); //Copy the exel file in Debug folder
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = true;
            string fileName = d + S;
            Workbook workbook = _excelApp.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];
            Range excelRange = worksheet.UsedRange;
            object[,] valueArray = (object[,])excelRange.get_Value(
                        XlRangeValueDataType.xlRangeValueDefault);
            RowNum = valueArray.GetLength(0);
            ColNum = valueArray.GetLength(1);
            string[] ar = new string[RowNum + 1];
            for (int i = 1; i <= RowNum; i++)
                ar[i] = (valueArray[i, 1].ToString());
            string[][] ar1 = new string[RowNum + 1][];
            for (int i = 1; i <= RowNum; i++)
            {
                string v = ar[i];
                ar1[i] = v.Split(',');
            }
            int[,] arr = new int[RowNum + 1, 4];
            for (int i = 1; i <= RowNum; i++)
                for (int j = 0; j < 3; j++)
                    arr[i, j + 1] = int.Parse(ar1[i][j]);
            try { workbook.Close(false, Type.Missing, Type.Missing); }
            catch { }
            try { _excelApp.Quit(); }
            catch { }
            return arr;
        }
        #endregion
        static double Dist(Object o1, Object o2)
        {
            double dR = o2.R - o1.R;
            double dG = o2.G - o1.G;
            double dB = o2.B - o1.B;
            double dist = Math.Sqrt(dR * dR + dG * dG + dB * dB);
            return dist;
        }
        static setOfObjects NeighborQuery(setOfObjects SetOfObjects, Object Object)
        {
            setOfObjects neighbors = new setOfObjects();
            for (int i = 0; i < SetOfObjects.Count; i++)
            {
                double dist = Dist(Object, SetOfObjects[i]);
                if (dist <= Eps)
                    neighbors.Add(SetOfObjects[i]);
            }
            return neighbors;
        }
        static double setCoreDistance(setOfObjects neighbors, Object Object)
        {
            if (neighbors.Count < MinPts)
                return -1;
            else
            {
                List<double> Distances = new List<double>();
                for (int i = 0; i < neighbors.Count; i++)
                    Distances.Add(Dist(Object, neighbors[i]));
                Distances.Sort();
                return Distances[MinPts - 1];/// MinPts omin hamsaye nazdik (ba khode Object hesab shode)
            }

        }
        static Queue<Object> updateOrderSeeds(Queue<Object> OrderSeeds, setOfObjects neighbors, Object CenterObject)
        {
            double c_dist = CenterObject.core_distance;
            foreach (Object Object in neighbors)
                if (!Object.Processed)
                {
                    double new_r_dist;
                    if (c_dist > Dist(CenterObject, Object))
                        new_r_dist = c_dist;
                    else
                        new_r_dist = Dist(CenterObject, Object);
                    if (Object.reachability_distance == -1)
                    {
                        Object.reachability_distance = new_r_dist;
                        Object.reachability_distance = new_r_dist;
                        OrderSeeds.Enqueue(Object);
                    }
                    else if (new_r_dist < Object.reachability_distance)
                    {
                        Object.reachability_distance = new_r_dist;
                        Queue<Object> Temp = new Queue<Object>(OrderSeeds.OrderBy(p => p.reachability_distance));
                        OrderSeeds.Clear();
                        OrderSeeds = Temp;
                    }
                }
            return OrderSeeds;
        }
        static setOfObjects ExpandClusterOrder(setOfObjects SetOfObjects, Object Object, setOfObjects OrderedFile)
        {
            setOfObjects neighbors = NeighborQuery(SetOfObjects, Object);
            Object.Processed = true;
            Object.reachability_distance = -1;
            Object.core_distance = setCoreDistance(neighbors, Object);  // core 0 shode ye mosht node ruye ham hastan
            OrderedFile.Add(Object);
            Queue<Object> OrderSeeds = new Queue<Object>();
            if (Object.core_distance != -1)
            {
                OrderSeeds = updateOrderSeeds(OrderSeeds, neighbors, Object);
                while (OrderSeeds.Count > 0)
                {
                    Object currentObject = OrderSeeds.Dequeue();
                    setOfObjects currentneighbors = NeighborQuery(SetOfObjects, currentObject);
                    currentObject.Processed = true;
                    currentObject.core_distance = setCoreDistance(currentneighbors, currentObject);
                    OrderedFile.Add(currentObject);
                    if (currentObject.core_distance != -1)
                        OrderSeeds = updateOrderSeeds(OrderSeeds, currentneighbors, currentObject);
                }
            }
            return OrderedFile;
        }
        static void ExtractDBSCANClustering(setOfObjects OrderedFile)
        {
            int ClusterId = 0;
            for (int i = 0; i < OrderedFile.Count; i++)
            {
                Object Object = OrderedFile[i];
                if (Object.reachability_distance > secondEps || Object.reachability_distance == -1)
                    if (Object.core_distance <= secondEps && Object.core_distance != -1)
                    {
                        ClusterId = ClusterId + 1;
                        Object.ClusterId = ClusterId;
                    }
                    else
                        Object.ClusterId = -1; //NOISE
                else
                    Object.ClusterId = ClusterId;
            }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            A = Read("\\HW2_data.xlsx");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            button1.Enabled = false;
            MinPts = Convert.ToInt16(textBox1.Text);
            Eps = Convert.ToDouble(textBox2.Text);
            secondEps = Convert.ToDouble(textBox3.Text);
            setOfObjects SetOfObjects = new setOfObjects();
            for (int i = 1; i <= RowNum; i++)
                SetOfObjects.Add(new Object(A[i, 1]/255.0, A[i, 2]/255.0, A[i, 3]/255.0, i));
            setOfObjects OrderedFile = new setOfObjects();
            for (int i = 0; i < SetOfObjects.Count; i++)
            {
                Object Object = SetOfObjects[i];
                if (!Object.Processed)
                    OrderedFile = ExpandClusterOrder(SetOfObjects, Object, OrderedFile);
            }
            // OrderedFile.Clear();
            int num = 1;
            foreach (Object O in OrderedFile)
            {
                listBox1.Items.Add(string.Format("{0}.           {1}", num, O.Name));
                num++;
            }

            ExtractDBSCANClustering(OrderedFile);

            string d = Directory.GetCurrentDirectory();
            chart1.ChartAreas[0].Axes[0].Title = "Points";
            chart1.ChartAreas[0].Axes[1].Title = "Reachability Distance";
            chart1.Series[0].Color = Color.Red;
            chart1.Series[0].BorderWidth = 1;
            chart1.ForeColor = Color.AliceBlue;
            chart1.Series[0].ChartType = SeriesChartType.Column;

            for (int i = 1; i < OrderedFile.Count; i++)
            {
                if (OrderedFile[i - 1].reachability_distance == -1)
                    chart1.Series[0].Points.AddXY(i, Eps);
                else
                    chart1.Series[0].Points.AddXY(i, OrderedFile[i - 1].reachability_distance);
                System.Windows.Forms.Application.DoEvents();
            }
            this.chart1.SaveImage(d + "\\chart.png", ChartImageFormat.Png);

            List<setOfObjects> Clusters = new List<setOfObjects>();
            int maxClusterId = OrderedFile.OrderBy(p => p.ClusterId).Last().ClusterId;
            if (maxClusterId <= 0)
                label2.Text = string.Format("All Points Are Noise!");
            else
                for (int i = 0; i < maxClusterId; i++)
                    Clusters.Add(new setOfObjects());
            foreach (Object O in SetOfObjects)
            {
                if (O.ClusterId > 0)
                    Clusters[O.ClusterId - 1].Add(O);
            }
            label3.Text = string.Format("THe Number Of Clusters Are : {0}", Clusters.Count);
            int sum = 0;
            for (int i = 0; i < Clusters.Count; i++)
            {
                sum += Clusters[i].Count;
                listBox2.Items.Add(string.Format("Cluster {0}    number of points : {1}", i + 1, Clusters[i].Count));
            }
            label2.Text = string.Format("{0} Points are Noise.", RowNum - sum);

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}

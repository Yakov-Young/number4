using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;
using Application = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Interop.Excel;

namespace number4
{
    public partial class Form1 : Form
    {

        List<double> minTime = new List<double>(); //bestTime
        List<Task> tskList = new List<Task>();
        List<int> arrBubble = new List<int>();
        List<int> arrInsert = new List<int>();
        List<int> arrShaker = new List<int>();
        List<int> arrQuick = new List<int>();
        List<int> arrBOGO = new List<int>();
        int allSort = 0, nonSort=0;
        private Stopwatch bubleTimes = new Stopwatch();  //Измерение затрат времени
        private Stopwatch insertTimes = new Stopwatch();
        private Stopwatch shakerTimes = new Stopwatch();
        private Stopwatch quickTimes = new Stopwatch();
        private Stopwatch BOGOTimes = new Stopwatch();
        private bool decrease;
        int x = -300, y= 0;
        

        public Form1()
        {
            InitializeComponent();
        }

        private void bubleSort(ZedGraphControl graphicObj)
        {

            bubleTimes.Restart();
            bubleTimes.Start();
            if (!decrease)
            {
                int temp;
                for (int i = 0; i < arrBubble.Count; i++)
                {
                    for (int j = i + 1; j < arrBubble.Count; j++)
                    {
                        if (arrBubble[i] > arrBubble[j])
                        {
                            temp = arrBubble[i];
                            arrBubble[i] = arrBubble[j];
                            arrBubble[j] = temp;
                        }

                        graphicSort(graphicObj, arrBubble);
                        Thread.Sleep(5);
                    }
                }
            }
            else
            {
                int temp;
                for (int i = 0; i < arrBubble.Count; i++)
                {
                    for (int j = i + 1; j < arrBubble.Count; j++)
                    {
                        if (arrBubble[i] < arrBubble[j])
                        {
                            temp = arrBubble[i];
                            arrBubble[i] = arrBubble[j];
                            arrBubble[j] = temp;
                        }

                        graphicSort(graphicObj, arrBubble);
                        Thread.Sleep(5);
                    }
                }
            }

            bubleTimes.Stop();
            minTime.Add(Math.Round((bubleTimes.Elapsed.TotalMilliseconds / 1000), 2));
            bubleTime.Invoke((MethodInvoker)delegate { bubleTime.Text = "bubleTime: " + Math.Round((bubleTimes.Elapsed.TotalMilliseconds / 1000), 2).ToString() + "s"; });

            allSort++;
            allSortCheck();

        }

        private void InsertingSort(ZedGraphControl graphicObj)
        {
            insertTimes.Restart();
            insertTimes.Start();
            if (!decrease)
            {
                for (var i = 1; i < arrInsert.Count; i++)
                {
                    var key = arrInsert[i];
                    var j = i;
                    while ((j > 0) && (arrInsert[j - 1] > key))
                    {

                        Exchange(arrInsert, j, j - 1);
                        j--;
                        graphicSort(graphicObj, arrInsert);
                        Thread.Sleep(5);

                    }

                    arrInsert[j] = key;

                }
            }
            else
            {
                for (var i = 1; i < arrInsert.Count; i++)
                {
                    var key = arrInsert[i];
                    var j = i;
                    while ((j > 0) && (arrInsert[j - 1] < key))
                    {

                        Exchange(arrInsert, j, j - 1);
                        j--;
                        graphicSort(graphicObj, arrInsert);
                        Thread.Sleep(5);

                    }

                    arrInsert[j] = key;

                }
            }
            
            insertTimes.Stop();
            minTime.Add(Math.Round((insertTimes.Elapsed.TotalMilliseconds / 1000), 2));
            insertTime.Invoke((MethodInvoker)delegate { insertTime.Text = "insertTime: " + Math.Round((insertTimes.Elapsed.TotalMilliseconds / 1000), 2).ToString() + "s"; });

            allSort++;
            allSortCheck();

        }

        private void ShakerSort(ZedGraphControl graphicObj)
        {
            shakerTimes.Restart();
            shakerTimes.Start();

            if (!decrease)
            {

                for (var i = 0; i < arrShaker.Count / 2; i++)
                {

                    var swapFlag = false;
                    //проход слева направо
                    for (var j = i; j < arrShaker.Count - i - 1; j++)
                    {
                        if (arrShaker[j] > arrShaker[j + 1])
                        {
                            Exchange(arrShaker, j, j + 1);
                            swapFlag = true;
                        }
                        graphicSort(graphicObj, arrShaker);
                        Thread.Sleep(5);

                    }

                    //проход справа налево
                    for (var j = arrShaker.Count - 2 - i; j > i; j--)
                    {
                        if (arrShaker[j - 1] > arrShaker[j])
                        {
                            Exchange(arrShaker, j - 1, j);
                            swapFlag = true;
                        }
                        graphicSort(graphicObj, arrShaker);
                        Thread.Sleep(5);
                    }

                    //если обменов не было выходим
                    if (!swapFlag)
                    {
                        break;
                    }

                }
            }
            else
            {

                for (var i = 0; i < arrShaker.Count / 2; i++)
                {

                    var swapFlag = false;
                    //проход слева направо
                    for (var j = i; j < arrShaker.Count - i - 1; j++)
                    {
                        if (arrShaker[j] < arrShaker[j + 1])
                        {
                            Exchange(arrShaker, j, j + 1);
                            swapFlag = true;
                        }
                        graphicSort(graphicObj, arrShaker);
                        Thread.Sleep(5);

                    }

                    //проход справа налево
                    for (var j = arrShaker.Count - 2 - i; j > i; j--)
                    {
                        if (arrShaker[j - 1] < arrShaker[j])
                        {
                            Exchange(arrShaker, j - 1, j);
                            swapFlag = true;
                        }
                        graphicSort(graphicObj, arrShaker);
                        Thread.Sleep(5);
                    }

                    //если обменов не было выходим
                    if (!swapFlag)
                    {
                        break;
                    }

                }
            }

            shakerTimes.Stop();
            minTime.Add(Math.Round((shakerTimes.Elapsed.TotalMilliseconds / 1000), 2));
            shakerTime.Invoke((MethodInvoker)delegate { shakerTime.Text = "shakerTime: " + Math.Round((shakerTimes.Elapsed.TotalMilliseconds / 1000), 2).ToString() + "s"; });

            allSort++;
            allSortCheck();
        }

        private int Partition(List<int> array, int minIndex, int maxIndex)
        {
            var pivot = minIndex - 1;
            if (!decrease)
            {
                for (var i = minIndex; i < maxIndex; i++)
                {
                    if (array[i] < array[maxIndex])
                    {
                        pivot++;
                        Exchange(array, pivot, i);
                    }
                    Thread.Sleep(5);
                }
            }
            else
            {
                for (var i = minIndex; i < maxIndex; i++)
                {
                    if (array[i] > array[maxIndex])
                    {
                        pivot++;
                        Exchange(array, pivot, i);
                    }
                    Thread.Sleep(5);
                }
            }

            pivot++;
            Exchange(arrQuick ,pivot,  maxIndex);
            return pivot;
        }

        //быстрая сортировка
        private void QuickSort(List<int> array, int minIndex, int maxIndex, ZedGraphControl graphicObj)
        {
            
            //quickTimes.Start();

            if (minIndex >= maxIndex)
            {
                
                return;
            }
            var pivotIndex = Partition(array, minIndex, maxIndex);
            graphicSort(graphicObj, arrQuick);
            
            QuickSort(array, minIndex, pivotIndex - 1, graphicObj);
            QuickSort(array, pivotIndex + 1, maxIndex, graphicObj);

            quickTimes.Stop();
            minTime.Add(Math.Round((quickTimes.Elapsed.TotalMilliseconds / 1000), 2));
            quickTime.Invoke((MethodInvoker)delegate { quickTime.Text = "quickTime: " + Math.Round((quickTimes.Elapsed.TotalMilliseconds / 1000), 2).ToString() + "s"; });
            
            allSort++;
            allSortCheck();
            
        }

       
        private void BogoSort(ZedGraphControl graphicObj)
        {
            BOGOTimes.Restart();
            BOGOTimes.Start();
            while (!IsSorted(arrBOGO))
            {
                Random random = new Random();
                var n = arrBOGO.Count;
                while (n > 1)
                {
                    n--;
                    var i = random.Next(n + 1);
                    Exchange(arrBOGO, i, n);

                }
                graphicSort(graphicObj, arrBOGO);
            }
            BOGOTimes.Stop();
            minTime.Add(Math.Round((BOGOTimes.Elapsed.TotalMilliseconds / 1000), 2));
            BOGOTime.Invoke((MethodInvoker)delegate { BOGOTime.Text = "BOGOTime: " + Math.Round((BOGOTimes.Elapsed.TotalMilliseconds / 1000), 2).ToString() + "s"; });
            allSort++;
            allSortCheck();

        }
         bool IsSorted(List<int> arr) //проверка сортирован ли список для BOGOsort
         {
            if (!decrease)
            {
                for (int i = 0; i < arr.Count - 1; i++)
                {
                    if (arr[i] > arr[i + 1])
                        return false;
                }

                return true;
            }
            else
            {
                for (int i = 0; i < arr.Count - 1; i++)
                {
                    if (arr[i] < arr[i + 1])
                        return false;
                }

                return true;
            }
         }

        private void allSortCheck() // bestSort
        {
            if(allSort == nonSort)  {  label2.Invoke((MethodInvoker)delegate{  label2.Text = minTime.Min().ToString();  });  }
        }
        private async void bubleAsync(ZedGraphControl graphicObj)
        {

            tskList.Add(Task.Run(() => bubleSort(graphicObj)));
            await Task.WhenAll(tskList);
            
        }

        private async void insertingAsync(ZedGraphControl graphicObj)
        {
            tskList.Add(Task.Run(() => InsertingSort(graphicObj)));
            await Task.WhenAll(tskList);
  
        }          

        private async void shakerAsync(ZedGraphControl graphicObj)
        {
            tskList.Add(Task.Run(() => ShakerSort(graphicObj)));
            await Task.WhenAll(tskList);
            
        }

        private async void quickAsync(ZedGraphControl graphicObj)
        {
            quickTimes.Restart();
            quickTimes.Start();
            Task task = Task.Run(() => QuickSort(arrQuick, 0, arrQuick.Count - 1, graphicObj));
            tskList.Add(task);
            
            await Task.WhenAll(task);
            
        }
        private async void bogoAsync(ZedGraphControl graphicObj)
        {
            tskList.Add(Task.Run(() => BogoSort(graphicObj)));
           
            await Task.WhenAll(tskList);

        }
        private void Exchange(List<int> arr, int i, int j)
        {
            var temp = arr[i];
            arr[i] = arr[j];
            arr[j] = temp;
        }
        private ZedGraphControl addGraph(string nameSort) //Выводит на экран сист. координат
        {
            ZedGraphControl graph = new ZedGraphControl();
            graph.Name = nameSort;
            graph.Width = 250;
            graph.Height = 200;
            x += 300;
            if (x >= groupBox1.Width)
            {
                x = 0;
                y += 250;
            }
            
            graph.Location = new System.Drawing.Point(x, y);


            groupBox1.Controls.Add(graph);


            return graph;
        }
        private void changeList() //Добавление элементов
        {

            if (arrBubble.Count > 0 || arrInsert.Count > 0 || arrShaker.Count > 0 
                                       || arrQuick.Count > 0 || arrBOGO.Count > 0)
            {
                arrBubble.Clear();
                arrInsert.Clear();
                arrShaker.Clear();
                arrQuick.Clear();
                arrBOGO.Clear();
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                int num = Convert.ToInt32(dataGridView1[0, i].Value);
                arrBubble.Add(num);
                arrInsert.Add(num);
                arrShaker.Add(num);
                arrQuick.Add(num);
                arrBOGO.Add(num);
            }
        }

        private void graphicSort(ZedGraphControl graphicObj, List<int> arr) //Воспроизведение изменений
        {

            GraphPane pane = graphicObj.GraphPane;
            pane.Title.Text = graphicObj.Name;
            pane.CurveList.Clear();
            int n = dataGridView1.Rows.Count - 1;
            double[] values = new double[n];
            for(int i=0; i < n; i++)
            {
                values[i] = arr[i];
            }
            pane.AddBar("Elem", null, values, Color.White);
            pane.BarSettings.MinClusterGap = 0F;

            graphicObj.AxisChange(); //обновление данных
            graphicObj.Invalidate(); //обновление графика

        }


        private void рассчитатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Controls.Clear();
            x = -300;
            y = 40;
            decrease = checkBox.Checked;


            if (checkBox1.Checked)
            {
                if (arrBubble.Count > 0 )
                {
                    ZedGraphControl graph = addGraph("BubleSort");
                    nonSort++;
                    bubleAsync(graph);
                }

            }
            if (checkBox2.Checked)
            {
                if (arrInsert.Count > 0 )
                {
                    ZedGraphControl graph = addGraph("InsertingSort");
                    nonSort++;
                    insertingAsync(graph);
                }

            }
            if (checkBox3.Checked)
            {
                if (arrShaker.Count > 0 )
                {
                    ZedGraphControl graph = addGraph("ShakerSort");
                    nonSort++;
                    shakerAsync(graph);
                }

            }
            if (checkBox4.Checked)
            {
                if (arrQuick.Count > 0 )
                {
                    ZedGraphControl graph = addGraph("QuickSort");
                    nonSort++;
                    quickAsync(graph);
                }

            }
            if (checkBox5.Checked)
            {
                if (arrBOGO.Count > 0 )
                {
                    ZedGraphControl graph = addGraph("BOGOSort");
                    nonSort++;
                    bogoAsync(graph);
                }

            }
        }


        private void генерацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                    int n = Convert.ToInt32(textBox1.Text);
                    Random rnd = new Random();
                    dataGridView1.Rows.Clear();
                    for (int i = 0; i < Convert.ToInt32(textBox1.Text); i++)
                    {
                        dataGridView1.Rows.Add();
                        dataGridView1[0, i].Value = rnd.Next(-50, 50);
                    }
                    changeList();

            }
            catch (Exception)
            {
                MessageBox.Show("Введите числовое значение");
            }
        }

        private void excelToolStripMenuItem1_Click(object sender, EventArgs e) //Получение данных из excel-файлов
        {
            openFileDialog1.FileName = String.Empty;
            DialogResult res = openFileDialog1.ShowDialog();
            if (res != DialogResult.OK) return;

            try
            {
                dataGridView1.Rows.Clear();
                Application ObjWorkExcel = new Application();
                Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
                if (ObjWorkSheet.Rows.CurrentRegion.EntireRow.Count == 1)
                {
                    MessageBox.Show("Excel файл пуст");
                }
                else
                {
                    dataGridView1.Rows.Clear();
                    var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
                    string num = String.Empty;
                    for (int i = 0; i < lastCell.Row; i++)
                    {
                        num = ObjWorkSheet.Cells[i + 1, 1].Text.ToString();
                        if (num.Trim() != String.Empty)
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1[0, i].Value = num;
                        }
                    }
                    changeList();
                    GC.Collect();
                }
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();

            }
            catch (Exception exception)
            {
                MessageBox.Show("При попытке загрузки из Excel произошла обшика!");
            }
        }


        private void льчиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            changeList();
            minTime.Clear();
            groupBox1.Controls.Clear();
        }

        private void textBox1_MouseClick(object sender, MouseEventArgs e)  //Отчистка поля генеации по нажтии не него
        {
            textBox1.Text = null;
            textBox1.ForeColor = Color.Black;
        }

    }

}

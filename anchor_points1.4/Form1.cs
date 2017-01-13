using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;                                 //для файлов
using Microsoft.JScript;
using Microsoft.JScript.Vsa;
using Microsoft;
using SpeechLib;
using System.Speech.Synthesis;

namespace anchor_points
{
    public partial class Form1 : Form
    {
        SpeechVoiceSpeakFlags Async = SpeechVoiceSpeakFlags.SVSFlagsAsync;
        SpeechVoiceSpeakFlags Sync = SpeechVoiceSpeakFlags.SVSFDefault;
        SpeechVoiceSpeakFlags Cancel = SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak;
        SpVoice speech = new SpVoice();

        DataGridView dgv=new DataGridView();

        string savepath;
        int colCount;

        public Form1()
        {
            InitializeComponent();
            List<string> strings = new List<string>();
            string[] s_points;//параметры как строки
            double[,] points; //параметры    
            int kol = 0; //количество строк в файле (и в матрице, собсна)
            int maxdeviation = 140;//"сектор"

            SpeechSynthesizer speaker = new SpeechSynthesizer();
            speech.Rate = 0;
            speech.Volume = 100;
            setvoice();

            /////////хш-хш

            opnfldlg.ShowDialog();

            savepath = Path.GetDirectoryName(opnfldlg.FileName);//куды потом сохранять            

            string fullpath = Path.GetFullPath(opnfldlg.FileName);
            tb1.Text = fullpath;
            FileStream file = new FileStream(fullpath, FileMode.Open);
            StreamReader reader = new StreamReader(file);

            while (reader.Peek() >= 0)
            {
                strings.Add(reader.ReadLine());  //строки из файла c рез. по маркерам в strings
                kol++;
            }
            reader.Close();
            file.Close();

            int n = 0;//количество столбцов, индекс идёт нафиг
            for (int i = 0; i < strings[0].Length - 1; i++) //определяем количество параметров (столбцов матрицы)
                if (strings[0][i] == ',')
                    n++;
            n += 2;//добавляются 2 параметра: cam_z и aberration

            points = new double[kol - 3, n];//kol-3 - верхняя строчка с подписями, первая и последняя - нули

            //здесь цикл  по всем strings, типа запись в матрицу points
            int q = 2;//строка с подписями не нужна, нулевая строка тож
            string[] sep = { ",", " " };
            object obj;

            while (q < strings.Count - 2)
            {
                s_points = strings[q].Split(sep, System.StringSplitOptions.RemoveEmptyEntries);
                for (int i = 1; i < s_points.Length; i++)//индексы не нужны
                {
                    obj = Microsoft.JScript.Eval.JScriptEvaluate(s_points[i], VsaEngine.CreateEngine());
                    points[q - 2, i - 1] = (double)obj;
                }
                q++;
            }

            //считывание aberrations и запись их в points
            strings.Clear();
            fullpath = Path.GetDirectoryName(opnfldlg.FileName);
            fullpath += "/aberrations.csv";
            file = new FileStream(fullpath, FileMode.Open);
            reader = new StreamReader(file);

            while (reader.Peek() >= 0)
            {
                strings.Add(reader.ReadLine());
            }
            reader.Close();
            file.Close();

            q = 1;//строка с подписями не нужна           
            while (q < strings.Count - 2)
            {
                s_points = strings[q].Split(sep, System.StringSplitOptions.RemoveEmptyEntries);
                obj = Microsoft.JScript.Eval.JScriptEvaluate(s_points[1], VsaEngine.CreateEngine());
                points[q - 1, n - 2] = (double)obj;
                q++;
            }

            //считывание cam_z и запись в points
            strings.Clear();
            fullpath = Path.GetDirectoryName(opnfldlg.FileName);
            fullpath += "/history.csv";
            file = new FileStream(fullpath, FileMode.Open);
            reader = new StreamReader(file);

            while (reader.Peek() >= 0)
            {
                strings.Add(reader.ReadLine());
            }
            reader.Close();
            file.Close();

            q = 1;//строка с подписями не нужна           
            while (q < strings.Count - 2)
            {
                s_points = strings[q].Split(sep, System.StringSplitOptions.RemoveEmptyEntries);
                obj = Microsoft.JScript.Eval.JScriptEvaluate(s_points[5], VsaEngine.CreateEngine());
                points[q - 1, n - 1] = (double)obj;
                q++;
            }

            //здесь создание матрицы с результатами
            int raz = (kol - 3) / 5;//кол-во строк новой матрицы
            if ((kol - 3) % 5 != 0)
                raz++;
            int[,] rezults = new int[raz, n + 1];//рез-ты по оп.точкам + 2 параметра + подитог (2 пар-ра уже учтены в n)
            int end = 4;         //5 шагов
            double sum = 0.0;//сумма aberration
            double[] sum2 = new double[(n - 1) / 2];//сумма deviations
            int ind = 0;//индекс для записи сумм deviations
            q = 0;
            bool flag = false;

            //общий цикл по строкам со смещением на 5
            #region all
            while (q < kol - 3)
            {
                int s = 5;//на сколько делить при нахождении средних, если шагов осталось меньше пяти
                if (q < kol - 3 - 4)//kol-3 - кол-во строк матрицы points
                    s = 5;
                else
                {
                    s = kol - 3 - q;
                    end = kol - 4;
                }
                if (end == q)
                    s = 1;
                //обработка cam_z
                for (int i = q; i < end; i++)
                {
                    if (points[i, n - 1] == points[i + 1, n - 1])
                        flag = false;//всё норм
                    else
                    {
                        flag = true;//было смещение
                        break;
                    }
                }
                if (flag == false)
                    rezults[q / 5, n - 1] = 0;
                else
                    rezults[q / 5, n - 1] = 1;
                flag = false;
                //обработка aberration и deviations
                for (int i = q; i <= end; i++)
                {
                    sum += points[i, n - 2];
                    for (int j = 1; j < n - 2; j += 2)
                    {
                        sum2[ind] += points[i, j];
                        ind++;
                    }
                    ind = 0;
                }
                //теперь в sum2 суммы deviation по каждой о.т.
                //aberration
                if (Math.Abs(sum) / s <= 40)
                    rezults[q / 5, n - 2] = 0;
                else
                    rezults[q / 5, n - 2] = 1;
                sum = 0.0;
                //deviations
                ind = 1;
                for (int i = 0; i < (n - 1) / 2; i++)
                {
                    if (Math.Abs(sum2[i]) / s <= maxdeviation)
                        rezults[q / 5, ind] = 0;
                    else
                        rezults[q / 5, ind] = 1;
                    ind += 2;
                }

                //обработка distances и определение № о.т., к кот. идёт движение, по наим. distance
                int yo = 0;//чётный/нечётный столбец  //нормальные герои всегда идут в обход)))
                int eta = 0;//№ текущей точки (0 - ни к какой точке)

                for (int j = 0; j < n - 3; j++)//только dist.
                {
                    if (yo == 0)//чётный, distance
                    {
                        eta++;
                        if (points[q, j] > points[end, j])//на 1-м шаге > на последнем
                        {
                            rezults[q / 5, j] = 0;//дистанция уменьшилась или не изменилась
                        }
                        else
                            rezults[q / 5, j] = 1;//дистанция увеличилась
                        //дополнительная проверка, если end=q
                        if (end == q)
                        {
                            if (points[q - 1, j] > points[end, j])//на предыдущем шаге > на текущем
                            {
                                rezults[q / 5, j] = 0;//дистанция уменьшилась или не изменилась
                            }
                            else
                                rezults[q / 5, j] = 1;//дистанция увеличилась
                        }

                        if (points[end, j] <= points[end, 0])//if distance текущей точки < min distance (<= на случай одной точки: номер должен записаться)
                        {
                            //проверка остальных параметров для этой точки и для данной итерации
                            if ((rezults[q / 5, n - 1] == 0) && (rezults[q / 5, n - 2] == 0) && (rezults[q / 5, j] == 0))
                            {
                                if (rezults[q / 5, j + 1] == 0)
                                    rezults[q / 5, n] = eta;
                            }
                        }
                        yo = 1;
                    }
                    else
                    {
                        yo = 0;
                    }
                }
                //обнулить sum2
                for (int i = 0; i < (n - 1) / 2; i++)
                    sum2[i] = 0.0;
                ind = 0;
                q += 5;
                end += 5;
            }
            #endregion all

            //общий результат
            double sum_r = 0.0, max = 0.0;//sum_r = сколько раз встречается точка
            int num = 0;//num - № о.т. с наиб. частотой
            int kol_t = (n - 2) / 2;//всего точек

            for (int j = 1; j <= kol_t; j++)//по всем точкам
            {
                for (int i = 0; i < raz; i++)//по всем строкам rezults
                {
                    if (rezults[i, n] == j)
                        sum_r++;
                }
                if (sum_r > max)
                {
                    max = sum_r;
                    num = j;
                }
                sum_r = 0.0;
            }

            sum_r = max / raz * 100;//вероятность движения

            //вывод результатов
            // dgv = new DataGridView()
            //{
            colCount = n + 3;
            dgv.ColumnCount = colCount;
            dgv.RowCount = kol-3;
            dgv.Font = new Font("Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dgv.Location = new Point(10, 20);
            dgv.AutoSize = false;
            dgv.Width = this.Width;
            dgv.Height = this.Height - 100;
            dgv.BackColor = System.Drawing.Color.AntiqueWhite;
            dgv.ScrollBars = ScrollBars.Both;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            dgv.EditMode = DataGridViewEditMode.EditProgrammatically;

            //};
            Controls.Add(dgv);
            dgv.CellEnter += new DataGridViewCellEventHandler(dgv_CellEnter);
            dgv.RowEnter += new DataGridViewCellEventHandler(dgv_CellEnter);
            dgv.CellFormatting += new DataGridViewCellFormattingEventHandler(dgv_CellFormatting);
            dgv.CellPainting += new DataGridViewCellPaintingEventHandler(dgv_CellPainting);

            foreach (DataGridViewColumn column in dgv.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            var label = new Label()
            {
                Location = new Point(10, 20 + dgv.Size.Height),
                Font = new Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204))),
                BackColor = System.Drawing.Color.LightGray,
                ForeColor = System.Drawing.Color.Black,
                AutoSize = true
            };
            Controls.Add(label);
            if (num == 0)
                label.Text = "В направлении движения нет ни одной опорной точки";

            dgv.Columns[0].HeaderText = "Индекс";
            dgv.Columns[1].HeaderText = "Шаг";

            int nn = 1, kuda = 3;
            while (kuda <= n - 1)
            {
                dgv.Columns[kuda].HeaderText = "В сектор 15' (т." + nn + ")";
                nn++;
                kuda += 2;
            }
            nn = 1; kuda = 2;
            while (kuda <= n - 2)
            {
                dgv.Columns[kuda].HeaderText = "Расстояние (до т." + nn + ")";
                nn++;
                kuda += 2;
            }

            dgv.Columns[n].HeaderText = "Приближение";
            dgv.Columns[n+1].HeaderText = "Маневр";
            dgv.Columns[n + 2].HeaderText = "№ точки";

            //индексы            
            for (int i = 0; i < kol - 3; i++)
                dgv[0, i].Value = i + 1;

            //записываем шаги            
            int step = 5;
            for (int i = 4; i < kol-3; i+=5)
            {
                dgv[1, i].Value = step;
                step += 5;
            }
//отсюда epta обозначает номер строки dgv, куда записывать
            int epta = 4;

            //сектор 15'  
            for (int i = 0; i < raz; i++)
            {
                int j = 1, k = 1;
                while (k <= kol_t)
                {
                    if (rezults[i, j] == 0)
                        dgv[j + 2, epta].Value = "Вписывается";
                    else
                        dgv[j + 2, epta].Value = "Не вписывается";
                    j += 2;
                    k++;
                }
                if (kol - 3 - epta <= 5)
                    epta = kol - 3 - 1;
                else
                    epta += 5;
            }
            //расстояние
            epta = 4;
            for (int i = 0; i < raz; i++)
            {
                int j = 0, k = 1;
                while (k <= kol_t)
                {
                    if (rezults[i, j] == 0)
                        dgv[j + 2, epta].Value = "Уменьшилось";
                    else
                        dgv[j + 2, epta].Value = "Не уменьшилось";
                    j += 2;
                    k++;
                }
                if (kol - 3 - epta <= 5)
                    epta = kol - 3 - 1;
                else
                    epta += 5;
            }
            //приближение
            epta = 4;
            for (int i = 0; i < raz; i++)
            {
                if (rezults[i, n - 1] == 0)
                    dgv[n, epta].Value = "Не было";
                else
                    dgv[n, epta].Value = "Было";
                if (kol - 3 - epta <= 5)
                    epta = kol - 3 - 1;
                else
                    epta += 5;
            }
            //маневр
            epta = 4;
            for (int i = 0; i < raz; i++)
            {
                if (rezults[i, n - 2] == 0)
                    dgv[n + 1, epta].Value = "Не было";
                else
                    dgv[n + 1, epta].Value = "Был";
                if (kol - 3 - epta <= 5)
                    epta = kol - 3 - 1;
                else
                    epta += 5;
            }
            //№ точки
            epta = 4;
            for (int i = 0; i < raz; i++)
            {
                dgv[n + 2, epta].Value = rezults[i, n];
                if (kol - 3 - epta <= 5)
                    epta = kol - 3 - 1;
                else
                    epta += 5;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Выберите файл с результатами по маркерам", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void setvoice()
        {
            FileStream file = new FileStream(Application.StartupPath + "\\SelectedVoice.txt", FileMode.Open);
            StreamReader reader = new StreamReader(file);
            string currvoicename = reader.ReadLine();
            reader.Close();
            file.Close();

            ISpeechObjectTokens Sotc = speech.GetVoices("", "");
            int n = 0;
            foreach (ISpeechObjectToken Sot in Sotc)
            {
                string tokenname = Sot.GetDescription(0);
                if (tokenname == currvoicename)
                {
                    speech.Voice = Sotc.Item(n);
                }
                n++;
            }
        }

        private void dgv_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            speech.Speak("",Cancel);
            if (dgv[colCount - 1, e.RowIndex].Value != null)
                speech.Speak("Двигаюсь к точке номер " + dgv[colCount-1, e.RowIndex].Value.ToString(), Async);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string s = "history";
            int ind = savepath.IndexOf(s) + 8;
            s = savepath.Substring(ind);
            ind = s.IndexOf('_');
            int ind2=s.IndexOf('.');
            string s2=s.Substring(ind+1,ind2-ind-1);
            s = s.Remove(ind);

            File.Delete(savepath + "\\" + s + "_anchor-points_" + s2 + ".csv");
            var sw = new StreamWriter(savepath+"\\"+s+"_anchor-points_"+s2+".csv", true, Encoding.Default);
            
            foreach (DataGridViewColumn column in dgv.Columns)//записываем шапку
            {
                sw.Write(column.HeaderText+";");
            }
            sw.WriteLine();
                       
            foreach (DataGridViewRow row in dgv.Rows)
                if (!row.IsNewRow)
                {
                    var first = true;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (!first) sw.Write(";");
                        if (cell.Value != null)
                            sw.Write(cell.Value.ToString());
                        else
                        {
                            //sw.WriteLine();
                            break;
                        }
                        first = false;
                    }
                    sw.WriteLine();
                }
            sw.Close(); 
        }

        private void dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

        }

        private void dgv_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {

        }
    }
}
//усё!

//opnfldlg.SafeFileName - имя ф-ла с расширением
// MessageBox.Show(s_points[0], "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

//foreach (String str in s_points)
//    tb1.Text += str + Environment.NewLine;

//points[q-1,i-1] = Convert.ToDouble(s_points[i]);

//просто вывод rezults в текстбокс
//int k = 1;
//for (int i = 0; i < (n - 2) / 2; i++)
//{
//    tb2.Text += "di" + k.ToString() + "  " + "de" + k.ToString() + "  ";
//    k++;
//}            
//tb2.Text += "abe  cam  №т." + Environment.NewLine; 
//for (int i = 0; i < raz; i++)
//{
//    for (int j = 0; j < n + 1; j++)
//        tb2.Text += rezults[i, j].ToString()+"      ";
//    tb2.Text += Environment.NewLine;
//}



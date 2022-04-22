using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Square
{
    public partial class Form1 : Form
    {
        SldWorks SwApp;
        IModelDoc2 swModel;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double K, L, M;
            K = Convert.ToDouble(textBox1.Text) / 1000;
            L = Convert.ToDouble(textBox2.Text) / 1000;
            M = Convert.ToDouble(textBox3.Text) / 1000;

            //int X, Y;
            //X = Convert.ToInt32(textBox6.Text);
            //Y = Convert.ToInt32(textBox7.Text);

            //Process[] processes = Process.GetProcessesByName("SLDWORKS");
            //foreach (Process process in processes)
            //{
            //    process.CloseMainWindow();
            //    process.Kill();
            //}
            //Guid myGuid1 = new Guid("F16137AD-8EE8-4D2A-8CAC-DFF5D1F67522");
            //object processSW = System.Activator.CreateInstance(System.Type.GetTypeFromCLSID(myGuid1));

            //SwApp = (SldWorks)processSW;
            //SwApp.Visible = true;

            try //блоком try-catch обычно оборачивают участок кода, который вызвает ошибку,   
            //и если она выскакивает, нужно ее отловить и распознать.  
            {
                SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                MessageBox.Show("Не удалось подключиться к solidworks");
                return;
            }
            if (SwApp.IActiveDoc == null)     //если никакой документ не открыт  
            {
                // создание 3d документа
                SwApp.NewPart();

                //MessageBox.Show("Надо открыть документ SW перед использованием");
                //System.Environment.Exit(-1); //моментальный выход из программы, если   
                //не выполняются наши условия -   
                //дальнейший код бесполезен.   
                //Только ошибки посыпятся  
            }

            swModel = SwApp.IActiveDoc2;

            // ====!!!! ТУТ НАЧИНАЕТСЯ МАГИЯ !!!!=====
            // ===== Построение основы =====
            swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0); //выбрал плоскость
            swModel.SketchManager.InsertSketch(true); //вставил эскиз в режиме редактирования 
            swModel.SketchManager.CreateCenterRectangle(0, 0, 0, L / 2, K / 2, 0);
            swModel.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, M, 0.01, false, false, false, false, 1, 1,
                false, false, false, false, true, true, true, 0, 0, false);

            // ======= Делаем вырез ========
            swModel.Extension.SelectByID2("", "FACE", L / 2, M / 2, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true); //вставил эскиз в режиме редактирования 
            swModel.SketchManager.CreateCornerRectangle(-K / 2 + 0.035, 0.050, 0, K / 2, M, 0); // 35 50
            swModel.FeatureManager.FeatureCut3(true, false, false, 1, 0, 0.01, 0.01, false, false, false, false, 0, 0,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            swModel.ClearSelection2(true);

            // ==== Наращиваем цилиндр =====
            //На этом моменте у меня уже начала ехать кукуха
            swModel.Extension.SelectByID2("", "FACE", 0, 0.050, - K / 2 + 0.001, false, 0, null, 0); // 50 35
            swModel.SketchManager.InsertSketch(true); //вставил эскиз в режиме редактирования 
            swModel.SketchManager.Create3PointArc(-0.045, -K / 2 + 0.035, 0, 0.045, -K / 2 + 0.035, 0, 0, -K / 2 + 0.035 + 0.045, 0); // 35 45r
            swModel.SketchManager.CreateLine(-0.045, -K / 2 + 0.035, 0, 0.045, -K / 2 + 0.035, 0); // 35, 45r
            swModel.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.040, 0.01, false, false, false, false, 1, 1,
                false, false, false, false, true, true, true, 0, 0, false); // 40
            //Тут я стал понимать, что я истинный гений

            // ==== Делаем цилиндрический вырез =====
            swModel.Extension.SelectByID2("", "FACE", 0, 0.001, K / 2, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true); //вставил эскиз в режиме редактирования
            swModel.SketchManager.Create3PointArc(-0.030, 0, 0, 0.030, 0, 0, 0, 0.030, 0); // 30r
            swModel.SketchManager.CreateLine(-0.030, 0, 0, 0.030, 0, 0); // 30r
            swModel.FeatureManager.FeatureCut3(true, false, false, 1, 0, 0.01, 0.01, false, false, false, false, 0, 0,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            swModel.ClearSelection2(true);

            // ======= Делаем вырез ========
            swModel.Extension.SelectByID2("", "FACE", L / 2, 0.001, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true); //вставил эскиз в режиме редактирования 
            swModel.SketchManager.CreateCornerRectangle(-0.040 / 2, 0.015, 0, 0.040 / 2, 0, 0); // 40 15
            swModel.FeatureManager.FeatureCut3(true, false, false, 1, 0, 0.01, 0.01, false, false, false, false, 0, 0,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            swModel.ClearSelection2(true);

            // ===== Делаем вырез еще ======
            swModel.Extension.SelectByID2("", "FACE", L / 2, 0.001, K / 2 - 0.001, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true); //вставил эскиз в режиме редактирования 
            swModel.SketchManager.CreateCornerRectangle(-K / 2, M, 0, -K / 2 + 0.020, M - 0.015, 0); // 20 15
            swModel.FeatureManager.FeatureCut3(true, false, false, 1, 0, 0.01, 0.01, false, false, false, false, 0, 0,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            swModel.ClearSelection2(true);

            // я настолько преисполнился в своем познании...

            //swModel.ClearSelection2(true);
            //swModel.SketchManager.InsertSketch(false);

            //swModel.Extension.SelectByID2("Эскиз2", "SKETCH", 0, 0, 0, false, 4, null, 0);
            //swModel.FeatureManager.FeatureCut3(true, false, false, 1, 0, 0.01, 0.01, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02,
            //    false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            //swModel.ClearSelection2(true);

            //swModel.Extension.SelectByID2("Вырез-Вытянуть1", "SOLIDBODY", 0.06, 0, 0.015, true, 256, null, 0);
            //swModel.Extension.SelectByID2("", "EDGE", -2.69215591075636E-02, -6.02367941999091E-02, 0.030334877568805, true, 1, null, 0);
            //swModel.Extension.SelectByID2("", "EDGE", -5.96297366384988E-02, -2.45181414630338E-02, 2.96297366384124E-02, true, 2, null, 0);

            //swModel.FeatureManager.FeatureLinearPattern4(X, K, Y, L, false, true, "NULL", "NULL", false, false, false, false, false, false, true, true, false, false, 0, 0);

            // построение шара
            //swModel.SketchManager.CreateCircle(0, 0, 0, -0.026781, -0.034413, 0);
            //swModel.SketchManager.CreateCenterLine(0, 0.080744, 0, 0, -0.074584, 0);
            //swModel.Extension.SelectByID2("Дуга1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            //swModel.SketchManager.SketchTrim(1, 4.92765597546085E-02, 4.95443671445792E-03, 0);
            //swModel.SketchManager.InsertSketch(true);
            //swModel.Extension.SelectByID2("Line1@Эскиз1", "EXTSKETCHSEGMENT", -0.289231981168354, 0.115291081382386, 0, true, 0, null, 0);            
        }
    }
}

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
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;

namespace Model2D
{
    public partial class Form1 : Form
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private SketchManager swSketchManager;
        private SelectionMgr swSelMgr;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Если не получилось открыть деталь в SolidWorks то завершить
            if (!GetSolidworks())
            {
                return;
            }

            double x0 = 100.0 / 1000.0;
            double y0 = 120.0 / 1000.0;

            double a, b;
            a = Convert.ToDouble(textBox1.Text) / 1000;
            b = Convert.ToDouble(textBox2.Text) / 1000;
            var line = swModel.SketchManager.CreateLine(x0, y0, 0, x0, y0 + b, 0).GetSketch().GetSketchPoints2();
            line = swModel.SketchManager.CreateLine(
                line[1].X,
                line[1].Y, 0,
                line[1].X + 0.020,
                line[1].Y, 0).GetSketch().GetSketchPoints2();
            line = swModel.SketchManager.CreateLine(
                line[2].X,
                line[2].Y, 0,
                line[2].X,
                line[2].Y - 0.035, 0).GetSketch().GetSketchPoints2();
            line = swModel.SketchManager.CreateLine(
                line[3].X,
                line[3].Y, 0,
                line[3].X + 0.020,
                line[3].Y, 0).GetSketch().GetSketchPoints2();
            line = swModel.SketchManager.CreateLine(
                line[4].X,
                line[4].Y, 0,
                line[4].X,
                line[4].Y + 0.035, 0).GetSketch().GetSketchPoints2();
            swModel.SketchManager.CreateLine(
                line[5].X,
                line[5].Y, 0,
                line[5].X + 1,
                line[5].Y, 0).GetSketch().GetSketchPoints2();

            swModel.SketchManager.CreateCircleByRadius(x0 + a, y0 + b - 0.020, 0, 0.015);
            swModel.SketchManager.CreateArc(x0 + a, y0 + b - 0.020, 0, x0 + a, y0 + b - 0.020 + 0.025, 0, x0 + a, y0 + b - 0.020 - 0.025, 0, -1);
            swModel.SketchManager.CreateLine(x0 + a - 0.015, y0 + b - 0.040, 0, x0 + 1, y0 + b - 0.040, 0);
            swModel.SketchManager.CreateCircleByRadius(x0 + 0.020, y0, 0, 0.015);

            swModel.SketchManager.CreateArc(x0 + 0.020, y0, 0, x0 + 0.020 - 0.020, y0, 0, x0 + 0.020 + 0.020, y0, 0, 1);
            swModel.SketchManager.CreateLine(x0 + a - 0.015, y0 + b - 0.040, 0, x0 + 0.020, y0 - 0.020, 0);

            swModel.Extension.SelectByID2("Line7", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("Arc4", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgFIXED");
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("Arc4", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("Line8", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgTANGENT");
            swModel.ClearSelection2(true);
            //swModel.Extension.SelectByID2("Line8", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            //swModel.SketchConstraintsDel(0, "sgFIXED");
            //swModel.ClearSelection2(true);

            swModel.Extension.SelectByID2("Line8", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            swModel.Extension.SelectByID2("Arc4", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.SketchTrim(0, x0 + 0.020, y0 + 0.020, 0);
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            swModel.Extension.SelectByID2("Arc2", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.SketchTrim(0, x0 + a, y0 + b, 0);
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("Line7", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            swModel.Extension.SelectByID2("Arc2", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.SketchTrim(0, x0 + a, y0 + b - 0.040, 0);
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("Arc2", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchManager.SketchTrim(0, x0 + 1, y0 + b, 0);
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("Arc2", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("Line7", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchManager.SketchTrim(0, x0 + 1, y0 + b - 0.040, 0);
            swModel.ClearSelection2(true);
        }

        private Boolean GetSolidworks()
        {
            try
            {
                // Присваиваем переменной ссылку на запущенный solidworks (по названию)
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch
            {
                // Отображает окно сообщения с заданным текстом
                MessageBox.Show("Не удалось найти запущенный Solidworks!");
                return false;
            }

            if (swApp.ActiveDoc == null)
            {
                // Отображает окно сообщения с заданным текстом
                MessageBox.Show("Надо открыть деталь перед использованием");
                return false;
            }


            // Присваиваем переменной ссылку на открытый активный проект в  SolidWorks
            swModel = (ModelDoc2)swApp.ActiveDoc;

            // Получает ISketchManager объект, который позволяет получить доступ к процедурам эскиза
            swSketchManager = (SketchManager)swModel.SketchManager;

            // Получает ISelectionMgr объект для данного документа, что делает выбранный объект доступным
            swSelMgr = (SelectionMgr)swModel.SelectionManager;


            // Проверка на открытие именно детали в SolidWorks
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                string text = "Это работает только на уровне чертежа";
                // Отображает окно сообщения с заданным текстом
                MessageBox.Show(text, "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}

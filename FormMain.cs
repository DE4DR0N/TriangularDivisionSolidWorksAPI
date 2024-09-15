using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace lab7
{
    public partial class FormMain : Form
    {
        private const double changeUnit = 1000;
        private uint count, strtCount, step;
        private int[] iters;
        private Entity ent;
        private Feature feature;
        private Feature skFeat;
        private Feature footing;
        private string front = "Спереди", top = "Сверху", right = "Справа";
        private double length, width, height, Spc2, P1x, P1y, P1z, P7x, P7y, P7z;
        private bool res;
        private SketchManager skm;
        private SldWorks swApp;
        private IModelDoc2 swModel;
        private SelectionMgr swSelMgr;

        public FormMain()
        {
            InitializeComponent();
        }

        private void btnBuild_Click(object sender, EventArgs e)
        {
            if (ObtainVaribles() == false) return;
            checkDrawing();
            CreatePoints(P1x, P1y, P1z, P7x, P7y, P7z);
            Drawing();
            btnBuild.Enabled = false;
            btnClear.Enabled = true;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            //feature.Select(false);
            //Feature swSubFeature = feature.GetFirstSubFeature() as Feature;

            //swModel.EditDelete();
            //swModel.ClearSelection();
            //swSubFeature.Select(false);
            //swModel.EditDelete();
            SelectSketch();
            swModel.EditDelete();

            btnBuild.Enabled = true;
            btnClear.Enabled = false;
        }

        private void btnLab5_Click(object sender, EventArgs e)
        {
            FormLab5 formLab5 = new FormLab5();
            formLab5.ShowDialog();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            //swApp = new SldWorks();
            //swApp.Visible = true;
        }

        #region Features

        private Feature featureCutDepthBack(double depth)
        {
            return swModel.FeatureManager.FeatureCut2(true, true, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature featureCutDepthFront(double depth)
        {
            return swModel.FeatureManager.FeatureCut2(true, false, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature featureCutThrough()
        {
            return swModel.FeatureManager.FeatureCut2(true, false, false, (int)swEndConditions_e.swEndCondThroughAllBoth, (int)swEndConditions_e.swEndCondThroughAllBoth,
                0, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature featureExtrusionBack(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(true, false, true,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        private Feature featureExtrusionBoth(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(false, true, false,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth / 2, depth / 2, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        private Feature featureExtrusionFront(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(true, false, false,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        #endregion Features

        #region Methods

        #region Obtaing

        private void checkDrawing()
        {
            //if (swApp.ActiveDoc == null)
            //{
            //    swModel = (ModelDoc2)swApp.INewPart();
            //    swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
            //    skm = swModel.SketchManager;
            //}
            //else
            //{
            //    swModel = (ModelDoc2)swApp.ActiveDoc;
            //    swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
            //    skm = swModel.SketchManager;
            //}
            if (!TryGetSolidworksApp(out this.swApp))
            {
                MessageBox.Show("SolidWorks не запущен");
                return;
            }

            if (!TryGetSolidWorksDocument(this.swApp, out this.swModel))
            {
                MessageBox.Show("Проект не открыт либо выбран неправильный тип документа");
                return;
            }

            if (!TryGetSolidWorksSketchManager(this.swModel, out this.skm))
            {
                MessageBox.Show("Не удается получить доступ к эскизу");
                return;
            }
        }
        private bool TryGetSolidworksApp(out SldWorks sw)
        {
            // Присваиваем переменной ссылку на запущенный solidworks (по названию)
            sw = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            return sw != null;
        }

        private bool TryGetSolidWorksDocument(SldWorks app, out IModelDoc2 doc)
        {
            // Присваиваем переменной ссылку на открытый активный проект в  SolidWorks
            doc = (ModelDoc2)app.IActiveDoc2;

            if (doc == null)
            {
                Console.WriteLine("null doc");
                return false;
            }
            if (doc.GetType() != (int)swDocumentTypes_e.swDocPART)
            {
                Console.WriteLine("not type");
                return false;
            }
            return true;
        }
        private bool TryGetSolidWorksSketchManager(IModelDoc2 doc, out SketchManager skMan)
        {
            // Получает ISketchManager объект, который позволяет получить доступ к процедурам эскиза
            skMan = (SketchManager)doc.SketchManager;
            return skMan != null;
        }

        private bool TryGetFeatureManager(IModelDoc2 swModel, out FeatureManager swFeatureManager)
        {
            swFeatureManager = swModel.FeatureManager;
            return swFeatureManager != null;
        }

        private bool ObtainVaribles()
        {
            try
            {
                Spc2 = Convert.ToDouble(textBox5.Text);
                Spc2 /= changeUnit;

                if (Spc2 < 0) throw new ArgumentException("Отступ меньше нуля");

                count = Convert.ToUInt32(textBox6.Text);
                strtCount = Convert.ToUInt32(textBox7.Text);
                step = Convert.ToUInt32(textBox4.Text);

                if (strtCount == 0 || count == 0) throw new ArgumentException("Параметры итераций не могут быть равны нулю");
                if (strtCount > count) throw new ArgumentException("Начальная итерация не может быть больше конечной");

                P1x = Convert.ToDouble(textBoxP1x.Text);
                P1y = Convert.ToDouble(textBoxP1y.Text);
                P1z = Convert.ToDouble(textBoxP1z.Text);

                P7x = Convert.ToDouble(textBoxP7x.Text);
                P7y = Convert.ToDouble(textBoxP7y.Text);
                P7z = Convert.ToDouble(textBoxP7z.Text);

                if (P7z < P1z)
                {
                    P1x = Convert.ToDouble(textBoxP7x.Text);
                    P1z = Convert.ToDouble(textBoxP7z.Text);
                    P7x = Convert.ToDouble(textBoxP1x.Text);
                    P7z = Convert.ToDouble(textBoxP1z.Text);
                }

                P1x /= changeUnit;
                P1y /= changeUnit;
                P1z /= changeUnit;
                P7x /= changeUnit;
                P7y /= changeUnit;
                P7z /= changeUnit;

                length = P7x;
                textBox1.Text = Convert.ToString(length * changeUnit);
                width = Math.Abs(P7z - P1z);
                textBox2.Text = Convert.ToString(width * changeUnit);
                height = Math.Abs(P7y - P1y);
                textBox3.Text = Convert.ToString(height * changeUnit);

                //цент.точка нахождение
                double centerX = (P1x + P7x) / 2.0;
                double centerY = P7y;
                double centerZ = (P1z + P7z) / 2.0;

                dataGridView1.Rows.Clear();

                dataGridView1.Rows.Add(P1x * 1000, P1y * 1000, P1z * 1000); //P1
                dataGridView1.Rows.Add(P1x * 1000, P7y * 1000, P1z * 1000);
                dataGridView1.Rows.Add(P1x * 1000, P7y * 1000, P7z * 1000);
                dataGridView1.Rows.Add(P1x * 1000, P1y * 1000, P7z * 1000);

                dataGridView1.Rows.Add(P7x * 1000, P1y * 1000, P1z * 1000);
                dataGridView1.Rows.Add(P7x * 1000, P7y * 1000, P1z * 1000);
                dataGridView1.Rows.Add(P7x * 1000, P7y * 1000, P7z * 1000); //P7
                dataGridView1.Rows.Add(P7x * 1000, P1y * 1000, P7z * 1000);

                dataGridView1.RowHeadersDefaultCellStyle.NullValue = "";
                dataGridView1.Rows[0].HeaderCell.Value = "P1";
                dataGridView1.Rows[1].HeaderCell.Value = "P2";
                dataGridView1.Rows[2].HeaderCell.Value = "P3";
                dataGridView1.Rows[3].HeaderCell.Value = "P4";
                dataGridView1.Rows[4].HeaderCell.Value = "P5";
                dataGridView1.Rows[5].HeaderCell.Value = "P6";
                dataGridView1.Rows[6].HeaderCell.Value = "P7";
                dataGridView1.Rows[7].HeaderCell.Value = "P8";

                iters = new int[count + 1];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обработки данных.\n" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        #endregion Obtaing

        public void CreatePoints(double P1_x, double P1_y, double P1_z, double P7_x, double P7_y, double P7_z)
        {
            SelectPlane(top);

            // Создаем точки
            skm.Insert3DSketch(true);
            SketchPoint point1 = skm.CreatePoint(P1_x, P1_y, P1_z); // P1
            SketchPoint point2 = skm.CreatePoint(P1_x, P7_y, P1_z);
            SketchPoint point3 = skm.CreatePoint(P1_x, P7_y, P7_z);
            SketchPoint point4 = skm.CreatePoint(P1_x, P1_y, P7_z);
            SketchPoint point5 = skm.CreatePoint(P7_x, P1_y, P1_z);
            SketchPoint point6 = skm.CreatePoint(P7_x, P7_y, P1_z);
            SketchPoint point7 = skm.CreatePoint(P7_x, P7_y, P7_z); // P7
            SketchPoint point8 = skm.CreatePoint(P7_x, P1_y, P7_z);
            SketchPoint center = skm.CreatePoint((P1_x + P7_x) / 2, P7_y, (P1_z + P7_z) / 2);

            swModel.ClearSelection();
            skm.Insert3DSketch(true);
            swModel.ClearSelection();
            point7.Select(true);
            Thread.Sleep(500);
            point6.Select(true);
            Thread.Sleep(500);
            point2.Select(true);
            Thread.Sleep(500);
            swModel.CreatePlaneThru3Points();

            swModel.ClearSelection();
        }

        private void Drawing()
        {
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayAnnotations, true);

            for (int i = 1; i <= count; i++)
            {
                iters[i] = i + 1;
            }

            for (uint i = strtCount; i <= count; i += step)
            {
                SelectPlane();
                skm.InsertSketch(false);
                if (i % 2 == 0)
                {
                    evenTriangle(i);
                }
                else
                {
                    oddTriangle(i);
                }
                if (i == count || i + step > count || step == 0) break;

                feature.Select(false);
                Feature swSubFeature = feature.GetFirstSubFeature() as Feature;

                swModel.EditDelete();
                swModel.ClearSelection();
                swSubFeature.Select(false);
                swModel.EditDelete();
            }
            swModel.ClearSelection();
        }

        private void SelectPlane(string name)
        {
            swModel.Extension.SelectByID2(name, "PLANE", 0, 0, 0, false, 0, null, 0);
        }

        private void SelectPlane()
        {
            swModel.Extension.SelectByID2("", "PLANE", P7x, P7y, P1z, false, 0, null, 0);
        }

        private void SelectSketch()
        {
            swModel.Extension.SelectByID2("", "SKETCH", P1x, P7y, P7z, false, 0, null, 0);
        }

        private void oddTriangle(uint count)
        {
            double xMax = length - Spc2;
            double xMin = P1x + Spc2;
            double yMax;
            if (P1z < 0) yMax = -(P1z + Spc2);
            else yMax = Math.Abs(P1z) - Spc2;

            switch (count)
            {
                case 1:
                    double lenMLine = length - 2 * Spc2;
                    double lMline = width - 2 * Spc2;

                    if (lMline < 2 / changeUnit || lenMLine < 2 / changeUnit)
                    {
                        MessageBox.Show("Построение невозможно. Уменьшите отступы", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        swModel.ClearSelection();
                        return;
                    }

                    var mLine = skm.CreateLine(xMin, yMax - lMline, 0, xMax, yMax - lMline, 0);
                    var line1 = skm.CreateLine(xMin, yMax - lMline, 0, (xMax + xMin) / 2, yMax, 0);
                    var line2 = skm.CreateLine(xMax, yMax - lMline, 0, (xMax + xMin) / 2, yMax, 0);

                    feature = featureCutDepthFront(height);

                    break;

                default:
                    lMline = (width - Spc2 * iters[count]) / ((count - 1) / 2);
                    double lKat = lMline / 2;

                    if (Math.Abs(lKat) < 1 / changeUnit || 3.62 * Spc2 >= Math.Abs(length))
                    {
                        MessageBox.Show("Построение невозможно. Уменьшите отступы", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        swModel.ClearSelection();
                        return;
                    }

                    var kat11 = skm.CreateLine(xMin, yMax, 0, xMax, yMax, 0);
                    var kat12 = skm.CreateLine(xMax, yMax, 0, xMax, yMax - lKat, 0);
                    var hip1 = skm.CreateLine(xMin, yMax, 0, xMax, yMax - lKat, 0);

                    for (int i = 1; i <= (count - 1) / 2; i++)
                    {
                        var mLineL = skm.CreateLine(xMin, yMax - Spc2, 0, xMin, yMax - Spc2 - lMline, 0);
                        var line1L = skm.CreateLine(xMin, yMax - Spc2, 0, xMax, yMax - lKat - Spc2, 0);
                        var line2L = skm.CreateLine(xMin, yMax - Spc2 - lMline, 0, xMax, yMax - lKat - Spc2, 0);

                        ModifyValue(ref yMax);
                        void ModifyValue(ref double val)
                        {
                            val = yMax - 2 * Spc2 - lMline;
                        }
                    }
                    if (P1z > 0) yMax = -(P1z + Spc2);
                    else yMax = Math.Abs(P1z) - Spc2;

                    for (int i = 1; i <= (count - 3) / 2; i++)
                    {
                        var mLineR = skm.CreateLine(xMax, yMax - 2 * Spc2 - lKat, 0, xMax, yMax - 2 * Spc2 - lKat - lMline, 0);
                        var line1R = skm.CreateLine(xMax, yMax - 2 * Spc2 - lKat, 0, xMin, yMax - 2 * Spc2 - lMline, 0);
                        var line2R = skm.CreateLine(xMax, yMax - 2 * Spc2 - lKat - lMline, 0, xMin, yMax - 2 * Spc2 - lMline, 0);

                        ModifyValue(ref yMax);
                        void ModifyValue(ref double val)
                        {
                            val = yMax - 2 * Spc2 - lMline;
                        }
                    }
                    yMax = yMax - 2 * Spc2 - lKat;

                    var kat21 = skm.CreateLine(xMin, yMax - lKat, 0, xMax, yMax - lKat, 0);
                    var kat22 = skm.CreateLine(xMax, yMax, 0, xMax, yMax - lKat, 0);
                    var hip2 = skm.CreateLine(xMax, yMax, 0, xMin, yMax - lKat, 0);

                    feature = featureCutDepthFront(height);
                    break;
            }
            return;
        }

        private void evenTriangle(uint count)
        {
            switch (count)
            {
                case 0:
                    break;

                default:
                    double xMax = length - Spc2;
                    double xMin = P1x + Spc2;
                    double yMax;
                    if (P1z > 0) yMax = -(P1z + Spc2);
                    else yMax = Math.Abs(P1z) - Spc2; ;

                    double lKat = (width - Spc2 * iters[count]) / (count / 2);

                    if (Math.Abs(lKat) < 1 / changeUnit || 3.62 * Spc2 >= Math.Abs(length))
                    {
                        MessageBox.Show("Построение невозможно. Уменьшите отступы", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        swModel.ClearSelection();

                        return;
                    }

                    for (uint i = count; i != 0; i -= 2)
                    {
                        var kat11 = skm.CreateLine(xMin, yMax, 0, xMax, yMax, 0);
                        var kat12 = skm.CreateLine(xMax, yMax, 0, xMax, yMax - lKat, 0);
                        var hip1 = skm.CreateLine(xMin, yMax, 0, xMax, yMax - lKat, 0);

                        var kat21 = skm.CreateLine(xMin, yMax - Spc2, 0, xMin, yMax - Spc2 - lKat, 0);
                        var kat22 = skm.CreateLine(xMin, yMax - Spc2 - lKat, 0, xMax, yMax - Spc2 - lKat, 0);
                        var hip2 = skm.CreateLine(xMin, yMax - Spc2, 0, xMax, yMax - Spc2 - lKat, 0);

                        ModifyValue(ref yMax);
                        void ModifyValue(ref double val)
                        {
                            val = yMax - 2 * Spc2 - lKat;
                        }
                    }
                    feature = featureCutDepthFront(height);
                    break;
            }
            return;
        }

        #endregion Methods
    }
}
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace lab7
{
    public partial class FormMain : Form
    {
        private const double ChangeUnit = 1000;
        private uint _count, _startCount, _step;
        private int[] iters;
        private Entity ent;
        private Feature feature;
        private Feature skFeat;
        private Feature footing;
        private const string FrontView = "Спереди", TopView = "Сверху", RightView = "Справа";
        private double length, width, height, Spc2, _p1X, _p1Y, _p1Z, _p7X, _p7Y, _p7Z;
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
            CheckDrawing();
            CreatePoints(_p1X, _p1Y, _p1Z, _p7X, _p7Y, _p7Z);
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

        #region Features

        private Feature FeatureCutDepthBack(double depth)
        {
            return swModel.FeatureManager.FeatureCut2(true, true, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature FeatureCutDepthFront(double depth)
        {
            return swModel.FeatureManager.FeatureCut2(true, false, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature FeatureCutThrough()
        {
            return swModel.FeatureManager.FeatureCut2(true, false, false, (int)swEndConditions_e.swEndCondThroughAllBoth, (int)swEndConditions_e.swEndCondThroughAllBoth,
                0, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature FeatureExtrusionBack(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(true, false, true,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        private Feature FeatureExtrusionBoth(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(false, true, false,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth / 2, depth / 2, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        private Feature FeatureExtrusionFront(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(true, false, false,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        #endregion Features

        #region Methods

        #region Obtaing

        private void CheckDrawing()
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

        private bool TryGetFeatureManager(IModelDoc2 doc, out FeatureManager swFeatureManager)
        {
            swFeatureManager = doc.FeatureManager;
            return swFeatureManager != null;
        }

        private bool ObtainVaribles()
        {
            try
            {
                Spc2 = Convert.ToDouble(textBox5.Text);
                Spc2 /= ChangeUnit;

                if (Spc2 < 0) throw new ArgumentException("Отступ меньше нуля");

                _count = Convert.ToUInt32(textBox6.Text);
                _startCount = Convert.ToUInt32(textBox7.Text);
                _step = Convert.ToUInt32(textBox4.Text);

                if (_startCount == 0 || _count == 0) throw new ArgumentException("Параметры итераций не могут быть равны нулю");
                if (_startCount > _count) throw new ArgumentException("Начальная итерация не может быть больше конечной");

                _p1X = Convert.ToDouble(textBoxP1x.Text);
                _p1Y = Convert.ToDouble(textBoxP1y.Text);
                _p1Z = Convert.ToDouble(textBoxP1z.Text);

                _p7X = Convert.ToDouble(textBoxP7x.Text);
                _p7Y = Convert.ToDouble(textBoxP7y.Text);
                _p7Z = Convert.ToDouble(textBoxP7z.Text);

                if (_p7Z < _p1Z)
                {
                    _p1X = Convert.ToDouble(textBoxP7x.Text);
                    _p1Z = Convert.ToDouble(textBoxP7z.Text);
                    _p7X = Convert.ToDouble(textBoxP1x.Text);
                    _p7Z = Convert.ToDouble(textBoxP1z.Text);
                }

                _p1X /= ChangeUnit;
                _p1Y /= ChangeUnit;
                _p1Z /= ChangeUnit;
                _p7X /= ChangeUnit;
                _p7Y /= ChangeUnit;
                _p7Z /= ChangeUnit;

                length = _p7X;
                textBox1.Text = (length * ChangeUnit).ToString();
                width = Math.Abs(_p7Z - _p1Z);
                textBox2.Text = (width * ChangeUnit).ToString();
                height = Math.Abs(_p7Y - _p1Y);
                textBox3.Text = (height * ChangeUnit).ToString();

                //цент.точка нахождение
                double centerX = (_p1X + _p7X) / 2.0;
                double centerY = _p7Y;
                double centerZ = (_p1Z + _p7Z) / 2.0;

                dataGridView1.Rows.Clear();

                dataGridView1.Rows.Add(_p1X * 1000, _p1Y * 1000, _p1Z * 1000); //P1
                dataGridView1.Rows.Add(_p1X * 1000, _p7Y * 1000, _p1Z * 1000);
                dataGridView1.Rows.Add(_p1X * 1000, _p7Y * 1000, _p7Z * 1000);
                dataGridView1.Rows.Add(_p1X * 1000, _p1Y * 1000, _p7Z * 1000);

                dataGridView1.Rows.Add(_p7X * 1000, _p1Y * 1000, _p1Z * 1000);
                dataGridView1.Rows.Add(_p7X * 1000, _p7Y * 1000, _p1Z * 1000);
                dataGridView1.Rows.Add(_p7X * 1000, _p7Y * 1000, _p7Z * 1000); //P7
                dataGridView1.Rows.Add(_p7X * 1000, _p1Y * 1000, _p7Z * 1000);

                dataGridView1.RowHeadersDefaultCellStyle.NullValue = "";
                dataGridView1.Rows[0].HeaderCell.Value = "P1";
                dataGridView1.Rows[1].HeaderCell.Value = "P2";
                dataGridView1.Rows[2].HeaderCell.Value = "P3";
                dataGridView1.Rows[3].HeaderCell.Value = "P4";
                dataGridView1.Rows[4].HeaderCell.Value = "P5";
                dataGridView1.Rows[5].HeaderCell.Value = "P6";
                dataGridView1.Rows[6].HeaderCell.Value = "P7";
                dataGridView1.Rows[7].HeaderCell.Value = "P8";

                iters = new int[_count + 1];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обработки данных.\n" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        #endregion Obtaing

        private void CreatePoints(double p1X, double p1Y, double p1Z, double p7X, double p7Y, double p7Z)
        {
            SelectPlane(TopView);

            // Создаем точки
            skm.Insert3DSketch(true);
            SketchPoint point1 = skm.CreatePoint(p1X, p1Y, p1Z); // P1
            SketchPoint point2 = skm.CreatePoint(p1X, p7Y, p1Z);
            SketchPoint point3 = skm.CreatePoint(p1X, p7Y, p7Z);
            SketchPoint point4 = skm.CreatePoint(p1X, p1Y, p7Z);
            SketchPoint point5 = skm.CreatePoint(p7X, p1Y, p1Z);
            SketchPoint point6 = skm.CreatePoint(p7X, p7Y, p1Z);
            SketchPoint point7 = skm.CreatePoint(p7X, p7Y, p7Z); // P7
            SketchPoint point8 = skm.CreatePoint(p7X, p1Y, p7Z);
            SketchPoint center = skm.CreatePoint((p1X + p7X) / 2, p7Y, (p1Z + p7Z) / 2);

            swModel.ClearSelection();
            skm.Insert3DSketch(true);
            swModel.ClearSelection();
            point7.Select(true);
            point6.Select(true);
            point2.Select(true);
            swModel.CreatePlaneThru3Points();

            swModel.ClearSelection();
        }

        private void Drawing()
        {
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayAnnotations, true);

            for (int i = 1; i <= _count; i++)
            {
                iters[i] = i + 1;
            }

            for (uint i = _startCount; i <= _count; i += _step)
            {
                SelectPlane();
                skm.InsertSketch(false);
                if (i % 2 == 0)
                {
                    EvenTriangle(i);
                }
                else
                {
                    OddTriangle(i);
                }
                if (i == _count || i + _step > _count || _step == 0) break;

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
            swModel.Extension.SelectByID2("", "PLANE", _p7X, _p7Y, _p1Z, false, 0, null, 0);
        }

        private void SelectSketch()
        {
            swModel.Extension.SelectByID2("", "SKETCH", _p1X, _p7Y, _p7Z, false, 0, null, 0);
        }

        private void OddTriangle(uint count)
        {
            double xMax = length - Spc2;
            double xMin = _p1X + Spc2;
            double yMax;
            if (_p1Z < 0) yMax = -(_p1Z + Spc2);
            else yMax = Math.Abs(_p1Z) - Spc2;

            switch (count)
            {
                case 1:
                    double lenMLine = length - 2 * Spc2;
                    double lMline = width - 2 * Spc2;

                    if (lMline < 2 / ChangeUnit || lenMLine < 2 / ChangeUnit)
                    {
                        MessageBox.Show("Построение невозможно. Уменьшите отступы", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        swModel.ClearSelection();
                        return;
                    }

                    var mLine = skm.CreateLine(xMin, yMax - lMline, 0, xMax, yMax - lMline, 0);
                    var line1 = skm.CreateLine(xMin, yMax - lMline, 0, (xMax + xMin) / 2, yMax, 0);
                    var line2 = skm.CreateLine(xMax, yMax - lMline, 0, (xMax + xMin) / 2, yMax, 0);

                    feature = FeatureCutDepthFront(height);

                    break;

                default:
                    lMline = (width - Spc2 * iters[count]) / ((count - 1) / 2);
                    double lKat = lMline / 2;

                    if (Math.Abs(lKat) < 1 / ChangeUnit || 3.62 * Spc2 >= Math.Abs(length))
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
                    if (_p1Z > 0) yMax = -(_p1Z + Spc2);
                    else yMax = Math.Abs(_p1Z) - Spc2;

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

                    feature = FeatureCutDepthFront(height);
                    break;
            }
        }

        private void EvenTriangle(uint count)
        {
            switch (count)
            {
                case 0:
                    break;

                default:
                    double xMax = length - Spc2;
                    double xMin = _p1X + Spc2;
                    double yMax;
                    if (_p1Z > 0) yMax = -(_p1Z + Spc2);
                    else yMax = Math.Abs(_p1Z) - Spc2; ;

                    double lKat = (width - Spc2 * iters[count]) / (count / 2);

                    if (Math.Abs(lKat) < 1 / ChangeUnit || 3.62 * Spc2 >= Math.Abs(length))
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
                    feature = FeatureCutDepthFront(height);
                    break;
            }
        }

        #endregion Methods
    }
}
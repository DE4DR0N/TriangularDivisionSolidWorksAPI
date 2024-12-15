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
        private double length, width, height, _marginInner, _marginOutX, _marginOutY, _p1X, _p1Y, _p1Z, _p7X, _p7Y, _p7Z;
        private SketchManager _skm;
        private SldWorks _swApp;
        private IModelDoc2 _swModel;
        private SelectionMgr _swSelMgr;

        public FormMain()
        {
            InitializeComponent();
        }

        private void btnBuild_Click(object sender, EventArgs e)
        {
            if (!ObtainVariables()) return;
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

            //_swModel.EditDelete();
            //_swModel.ClearSelection();
            //swSubFeature.Select(false);
            //_swModel.EditDelete();
            SelectSketch();
            _swModel.EditDelete();

            btnBuild.Enabled = true;
            btnClear.Enabled = false;
        }

        #region Features

        private Feature FeatureCutDepthBack(double depth)
        {
            return _swModel.FeatureManager.FeatureCut2(true, true, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature FeatureCutDepthFront(double depth)
        {
            return _swModel.FeatureManager.FeatureCut2(true, false, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature FeatureCutThrough()
        {
            return _swModel.FeatureManager.FeatureCut2(true, false, false, (int)swEndConditions_e.swEndCondThroughAllBoth, (int)swEndConditions_e.swEndCondThroughAllBoth,
                0, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature FeatureExtrusionBack(double depth)
        {
            return _swModel.FeatureManager.FeatureExtrusion2(true, false, true,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        private Feature FeatureExtrusionBoth(double depth)
        {
            return _swModel.FeatureManager.FeatureExtrusion2(false, true, false,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth / 2, depth / 2, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        private Feature FeatureExtrusionFront(double depth)
        {
            return _swModel.FeatureManager.FeatureExtrusion2(true, false, false,
                (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
                true, true, 0, 0, false);
        }

        #endregion Features

        #region Methods

        #region Obtaing

        private void CheckDrawing()
        {
            //if (_swApp.ActiveDoc == null)
            //{
            //    _swModel = (ModelDoc2)_swApp.INewPart();
            //    _swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
            //    _skm = _swModel.SketchManager;
            //}
            //else
            //{
            //    _swModel = (ModelDoc2)_swApp.ActiveDoc;
            //    _swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
            //    _skm = _swModel.SketchManager;
            //}
            if (!TryGetSolidworksApp(out this._swApp))
            {
                MessageBox.Show("SolidWorks не запущен");
                return;
            }

            if (!TryGetSolidWorksDocument(this._swApp, out this._swModel))
            {
                MessageBox.Show("Проект не открыт либо выбран неправильный тип документа");
                return;
            }

            if (!TryGetSolidWorksSketchManager(this._swModel, out this._skm))
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

        private bool ObtainVariables()
        {
            try
            {
                _marginInner = Convert.ToDouble(nmrcUpDownMargin.Value) / ChangeUnit;
                _marginOutX = Convert.ToDouble(nmrcUpDownMarginX.Value) / ChangeUnit;
                _marginOutY = Convert.ToDouble(nmrcUpDownMarginY.Value) / ChangeUnit;

                if (_marginInner < 0 || _marginOutX < 0 || _marginOutY < 0) throw new ArgumentException("Отступ меньше нуля");

                _count = Convert.ToUInt32(nmrcUpDownIters.Value);
                _startCount = Convert.ToUInt32(nmrcUpDownFirstIter.Value);
                _step = Convert.ToUInt32(nmrcUpDownStep.Value);

                if (_startCount == 0 || _count == 0) throw new ArgumentException("Параметры итераций не могут быть равны нулю");
                if (_startCount > _count) throw new ArgumentException("Начальная итерация не может быть больше конечной");

                if (nmrcUpDownP7z.Value < nmrcUpDownP1z.Value)
                {
                    _p1X = Convert.ToDouble(nmrcUpDownP7x.Value);
                    _p1Z = Convert.ToDouble(nmrcUpDownP7z.Value);
                    _p7X = Convert.ToDouble(nmrcUpDownP1x.Value);
                    _p7Z = Convert.ToDouble(nmrcUpDownP1z.Value);
                }
                else
                {
                    _p1X = Convert.ToDouble(nmrcUpDownP1x.Value);
                    _p1Z = Convert.ToDouble(nmrcUpDownP1z.Value);
                    _p7X = Convert.ToDouble(nmrcUpDownP7x.Value);
                    _p7Z = Convert.ToDouble(nmrcUpDownP7z.Value);
                }

                _p1Y = Convert.ToDouble(nmrcUpDownP1y.Value);
                _p7Y = Convert.ToDouble(nmrcUpDownP7y.Value);
                
                dataGridView1.Rows.Clear();

                dataGridView1.Rows.Add(_p1X, _p1Y, _p1Z); //P1
                dataGridView1.Rows.Add(_p1X, _p7Y, _p1Z);
                dataGridView1.Rows.Add(_p1X, _p7Y, _p7Z);
                dataGridView1.Rows.Add(_p1X, _p1Y, _p7Z);

                dataGridView1.Rows.Add(_p7X, _p1Y, _p1Z);
                dataGridView1.Rows.Add(_p7X, _p7Y, _p1Z);
                dataGridView1.Rows.Add(_p7X, _p7Y, _p7Z); //P7
                dataGridView1.Rows.Add(_p7X, _p1Y, _p7Z);

                dataGridView1.RowHeadersDefaultCellStyle.NullValue = "";
                dataGridView1.Rows[0].HeaderCell.Value = "P1";
                dataGridView1.Rows[1].HeaderCell.Value = "P2";
                dataGridView1.Rows[2].HeaderCell.Value = "P3";
                dataGridView1.Rows[3].HeaderCell.Value = "P4";
                dataGridView1.Rows[4].HeaderCell.Value = "P5";
                dataGridView1.Rows[5].HeaderCell.Value = "P6";
                dataGridView1.Rows[6].HeaderCell.Value = "P7";
                dataGridView1.Rows[7].HeaderCell.Value = "P8";
                
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
                var centerX = (_p1X + _p7X) / 2.0;
                var centerY = _p7Y;
                var centerZ = (_p1Z + _p7Z) / 2.0;

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
            // SelectPlane(TopView);

            // Создаем точки
            _skm.Insert3DSketch(true);
            SketchPoint point1 = _skm.CreatePoint(p1X, p1Y, p1Z); // P1
            SketchPoint point2 = _skm.CreatePoint(p1X, p7Y, p1Z);
            SketchPoint point3 = _skm.CreatePoint(p1X, p7Y, p7Z);
            SketchPoint point4 = _skm.CreatePoint(p1X, p1Y, p7Z);
            SketchPoint point5 = _skm.CreatePoint(p7X, p1Y, p1Z);
            SketchPoint point6 = _skm.CreatePoint(p7X, p7Y, p1Z);
            SketchPoint point7 = _skm.CreatePoint(p7X, p7Y, p7Z); // P7
            SketchPoint point8 = _skm.CreatePoint(p7X, p1Y, p7Z);
            SketchPoint center = _skm.CreatePoint((p1X + p7X) / 2, p7Y, (p1Z + p7Z) / 2);

            _swModel.ClearSelection();
            _skm.Insert3DSketch(true);
            _swModel.ClearSelection();
            point7.Select(true);
            point6.Select(true);
            point2.Select(true);
            _swModel.CreatePlaneThru3Points();

            _swModel.ClearSelection();
        }

        private void Drawing()
        {
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            _swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayAnnotations, true);

            for (int i = 1; i <= _count; i++)
            {
                iters[i] = i + 1;
            }

            for (uint i = _startCount; i <= _count; i += _step)
            {
                SelectPlane();
                _skm.InsertSketch(false);
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

                _swModel.EditDelete();
                _swModel.ClearSelection();
                swSubFeature.Select(false);
                _swModel.EditDelete();
            }
            _swModel.ClearSelection();
        }

        private void SelectPlane(string name)
        {
            _swModel.Extension.SelectByID2(name, "PLANE", 0, 0, 0, false, 0, null, 0);
        }

        private void SelectPlane()
        {
            _swModel.Extension.SelectByID2("", "PLANE", _p7X, _p7Y, _p1Z, false, 0, null, 0);
        }

        private void SelectSketch()
        {
            _swModel.Extension.SelectByID2("", "SKETCH", _p1X, _p7Y, _p7Z, false, 0, null, 0);
        }

        private void OddTriangle(uint count)
        {
            double xMax = length - _marginOutX;
            double xMin = _p1X + _marginOutX;
            double yMax;
            if (_p1Z < 0) yMax = -(_p1Z + _marginOutY);
            else yMax = Math.Abs(_p1Z) - _marginOutY;

            switch (count)
            {
                case 1:
                    double lenMLine = length - 2 * _marginOutX;
                    double heightLine = width - 2 * _marginOutY;

                    if (heightLine < 2 / ChangeUnit || lenMLine < 2 / ChangeUnit)
                    {
                        _swModel.ClearSelection();
                        SelectSketch();
                        _swModel.EditDelete();

                        btnBuild.Enabled = true;
                        btnClear.Enabled = false;
                        throw new ArgumentException("Неверные отступы");
                    }

                    var mLine = _skm.CreateLine(xMin, yMax - heightLine, 0, xMax, yMax - heightLine, 0);
                    var line1 = _skm.CreateLine(xMin, yMax - heightLine, 0, (xMax + xMin) / 2, yMax, 0);
                    var line2 = _skm.CreateLine(xMax, yMax - heightLine, 0, (xMax + xMin) / 2, yMax, 0);

                    feature = FeatureCutDepthFront(height);

                    break;

                default:
                    heightLine = (width - 2 * _marginOutY - _marginInner * (iters[count] - 2)) / ((count - 1) / 2);
                    double heightLeg = heightLine / 2;

                    if (Math.Abs(heightLeg) < 1 / ChangeUnit || 3.62 * _marginInner >= Math.Abs(length))
                    {
                        _swModel.ClearSelection();
                        SelectSketch();
                        _swModel.EditDelete();

                        btnBuild.Enabled = true;
                        btnClear.Enabled = false;
                        throw new ArgumentException("Неверные отступы");
                    }

                    var kat11 = _skm.CreateLine(xMin, yMax, 0, xMax, yMax, 0);
                    var kat12 = _skm.CreateLine(xMax, yMax, 0, xMax, yMax - heightLeg, 0);
                    var hip1 = _skm.CreateLine(xMin, yMax, 0, xMax, yMax - heightLeg, 0);

                    for (int i = 1; i <= (count - 1) / 2; i++)
                    {
                        var mLineL = _skm.CreateLine(xMin, yMax - _marginInner, 0, xMin, yMax - _marginInner - heightLine, 0);
                        var line1L = _skm.CreateLine(xMin, yMax - _marginInner, 0, xMax, yMax - heightLeg - _marginInner, 0);
                        var line2L = _skm.CreateLine(xMin, yMax - _marginInner - heightLine, 0, xMax, yMax - heightLeg - _marginInner, 0);

                        ModifyValue(ref yMax);
                        void ModifyValue(ref double val)
                        {
                            val = yMax - 2 * _marginInner - heightLine;
                        }
                    }
                    if (_p1Z > 0) yMax = -(_p1Z + _marginOutY);
                    else yMax = Math.Abs(_p1Z) - _marginOutY;

                    for (int i = 1; i <= (count - 3) / 2; i++)
                    {
                        var mLineR = _skm.CreateLine(xMax, yMax - 2 * _marginInner - heightLeg, 0, xMax, yMax - 2 * _marginInner - heightLeg - heightLine, 0);
                        var line1R = _skm.CreateLine(xMax, yMax - 2 * _marginInner - heightLeg, 0, xMin, yMax - 2 * _marginInner - heightLine, 0);
                        var line2R = _skm.CreateLine(xMax, yMax - 2 * _marginInner - heightLeg - heightLine, 0, xMin, yMax - 2 * _marginInner - heightLine, 0);

                        ModifyValue(ref yMax);
                        void ModifyValue(ref double val)
                        {
                            val = yMax - 2 * _marginInner - heightLine;
                        }
                    }
                    yMax = yMax - 2 * _marginInner - heightLeg;

                    var kat21 = _skm.CreateLine(xMin, yMax - heightLeg, 0, xMax, yMax - heightLeg, 0);
                    var kat22 = _skm.CreateLine(xMax, yMax, 0, xMax, yMax - heightLeg, 0);
                    var hip2 = _skm.CreateLine(xMax, yMax, 0, xMin, yMax - heightLeg, 0);

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
                    double xMax = length - _marginOutX;
                    double xMin = _p1X + _marginOutX;
                    double yMax;
                    if (_p1Z > 0) yMax = -(_p1Z + _marginOutY);
                    else yMax = Math.Abs(_p1Z) - _marginOutY;

                    double heightLeg = (width - 2 * _marginOutY - _marginInner * (iters[count] - 2)) / (count / 2);

                    if (Math.Abs(heightLeg) < 1 / ChangeUnit || 3.62 * _marginInner >= Math.Abs(length))
                    {
                        _swModel.ClearSelection();
                        SelectSketch();
                        _swModel.EditDelete();

                        btnBuild.Enabled = true;
                        btnClear.Enabled = false;
                        throw new ArgumentException("Неверные отступы");
                    }

                    for (uint i = count; i != 0; i -= 2)
                    {
                        var kat11 = _skm.CreateLine(xMin, yMax, 0, xMax, yMax, 0);
                        var kat12 = _skm.CreateLine(xMax, yMax, 0, xMax, yMax - heightLeg, 0);
                        var hip1 = _skm.CreateLine(xMin, yMax, 0, xMax, yMax - heightLeg, 0);

                        var kat21 = _skm.CreateLine(xMin, yMax - _marginInner, 0, xMin, yMax - _marginInner - heightLeg, 0);
                        var kat22 = _skm.CreateLine(xMin, yMax - _marginInner - heightLeg, 0, xMax, yMax - _marginInner - heightLeg, 0);
                        var hip2 = _skm.CreateLine(xMin, yMax - _marginInner, 0, xMax, yMax - _marginInner - heightLeg, 0);

                        ModifyValue(ref yMax);
                        void ModifyValue(ref double val)
                        {
                            val = yMax - 2 * _marginInner - heightLeg;
                        }
                    }
                    feature = FeatureCutDepthFront(height);
                    break;
            }
        }

        #endregion Methods
    }
}
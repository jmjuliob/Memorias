using ETABSv1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WFAAPIETABS
{
    class cFuncionesEtabs
    {
        public static cOAPI Etabs_App;
        public static cSapModel Modelo;
        public static cHelper MyHElper;


        public static void Open_Etabs()
        {

            Application.UseWaitCursor = true;
            string Ruta_Archivo;
            MyHElper = new Helper();

            try
            {
                Etabs_App = (cOAPI)Marshal.GetActiveObject("CSI.ETABS.API.ETABSObject");
                Modelo = Etabs_App.SapModel;
                Ruta_Archivo = Modelo.GetModelFilepath();
            }
            catch
            {
                Etabs_App = MyHElper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");
                Etabs_App.ApplicationStart();

                OpenFileDialog Myfile = new OpenFileDialog
                {
                    Filter = "Archivo de ETABS|*.edb",
                    Title = "Abrir archivo"
                };
                Myfile.ShowDialog();
                Ruta_Archivo = Myfile.FileName;
                Modelo = Etabs_App.SapModel;
                Modelo.File.OpenFile(Ruta_Archivo);
            }

            Modelo.SetPresentUnits(eUnits.Ton_m_C);
            Application.UseWaitCursor = false;
        }



        //public static void Get_Pier_Forces(ref Listas_Objetos Lista_i, string Piso, ref ProgressBar progressBar)
        //{
        //    #region Variables wall geometry

        //    int Numero_Elementos = 0;
        //    string[] Pier_Labels = null;
        //    string[] Pier_Labels_aux = null;
        //    string[] StoryName = null;
        //    double[] AxisAngle = null;
        //    int[] NumAreaObjs = null;
        //    int[] NumLineObjs = null;
        //    double[] WidthBot = null;
        //    double[] ThicknessBot = null;
        //    double[] WidthTop = null;
        //    double[] ThicknessTop = null;
        //    string[] MatProp = null;
        //    double[] CGBotX = null;
        //    double[] CGBotY = null;
        //    double[] CGBotZ = null;
        //    double[] CGTopX = null;
        //    double[] CGTopY = null;
        //    double[] CGTopZ = null;

        //    #endregion Variables wall geometry

        //    #region Variables_pier_forces

        //    string[] LoadCase = null;
        //    string[] Location = null;
        //    double[] P = null;
        //    double[] V2 = null;
        //    double[] V3 = null;
        //    double[] T = null;
        //    double[] M2 = null;
        //    double[] M3 = null;

        //    #endregion Variables_pier_forces

        //    int Inicio, fin;

        //    Muro Muro_i;
        //    if (Modelo.GetModelIsLocked() == false) Modelo.Analyze.RunAnalysis();

        //    List<string> Load_Cases = new List<string>();
        //    Load_Cases.AddRange(new string[] { "Dead", "Live", "A" });

        //    cPierLabel Muro = Modelo.PierLabel;
        //    cAnalysisResults Analisis = Modelo.Results;

        //    Set_Load_Cases(Load_Cases, ref Analisis, 0);
        //    Analisis.PierForce(ref Numero_Elementos, ref StoryName, ref Pier_Labels_aux, ref LoadCase, ref Location, ref P, ref V2, ref V3, ref T, ref M2, ref M3);


        //    Muro.GetNameList(ref Numero_Elementos, ref Pier_Labels);

        //    progressBar.Visible = true;
        //    progressBar.Value = 0;
        //    progressBar.Maximum = Pier_Labels.Count();

        //    foreach (string Pier_label in Pier_Labels)
        //    {
        //        Muro.GetSectionProperties(Pier_label, ref Numero_Elementos, ref StoryName, ref AxisAngle, ref NumAreaObjs, ref NumLineObjs, ref WidthBot, ref ThicknessBot,
        //            ref WidthTop, ref ThicknessTop, ref MatProp, ref CGBotX, ref CGBotY, ref CGBotZ, ref CGTopX, ref CGTopY, ref CGTopZ);

        //        if (ThicknessBot != null)
        //        {
        //            Inicio = Pier_Labels_aux.ToList().FindIndex(x => x == Pier_label);
        //            fin = Pier_Labels_aux.ToList().FindLastIndex(x => x == Pier_label);

        //            Muro_i = new Muro

        //            {
        //                Label = Pier_label,
        //                Bw = ThicknessBot.ToList().Select(x => Math.Round(x, 2)).ToList(),
        //                Materiales = MatProp.ToList(),
        //                lw = WidthTop.ToList().Select(x => Math.Round(x, 2)).ToList(),
        //                Story = Piso,
        //                Load_Cases = new List<string>(),
        //                P_load = new List<double>(),
        //                P_dist = new List<double>(),
        //                Shells = new List<Pier_Shell>()
        //            };

        //            if (Inicio >= 0)
        //            {
        //                for (int j = Inicio; j <= fin; j++)
        //                {
        //                    if (Location[j] == "Top")
        //                    {
        //                        Muro_i.Load_Cases.Add(LoadCase[j]);
        //                        Muro_i.P_load.Add(Math.Round(P[j], 3));
        //                        Muro_i.P_dist.Add(Math.Round(P[j] / Muro_i.lw.Last(), 3));
        //                    }
        //                }
        //                Lista_i.Lista_Muros.Add(Muro_i);
        //            }
        //        }
        //        progressBar.Increment(1);
        //    }
        //    progressBar.Visible = false;
        //    Get_area(Lista_i, Piso);
        //}

        //public static void SelectionPoints(ref List<double> CoordX, ref List<double> CoordY, ref List<double> CoordZ)
        //{
        //    int NumberSelec = 0;
        //    int[] ObjeType = { };
        //    string[] ObjeName = { };
        //    CoordX.Clear();
        //    CoordY.Clear();

        //    Modelo.SelectObj.GetSelected(ref NumberSelec, ref ObjeType, ref ObjeName);
        //    Modelo.SelectObj.ClearSelection();

        //    for (int i = 0; i < ObjeType.Length; i++)
        //    {
        //        if (ObjeType[i] == 1)
        //        {
        //            double X = 0; double Y = 0; double Z = 0;
        //            Modelo.PointObj.GetCoordCartesian(ObjeName[i], ref X, ref Y, ref Z);
        //            CoordX.Add(X);
        //            CoordY.Add(Y);
        //            CoordZ.Add(Z);
        //        }
        //    }
        //}

        //public static void DrawArea(List<double> CoordX, List<double> CoordY, List<double> CoordZ, string PropName)
        //{
        //    int NoPuntos = CoordX.Count;
        //    string Name = "";
        //    var X = CoordX.ToArray(); var Y = CoordY.ToArray(); var Z = CoordZ.ToArray();
        //    Modelo.AreaObj.AddByCoord(NoPuntos, ref X, ref Y, ref Z, ref Name, PropName);
        //    Modelo.View.RefreshView();
        //}

        //public static void GetSlabs(ref List<string> NamesSlabs)
        //{
        //    NamesSlabs.Clear();
        //    int number = 0; string[] names = { };
        //    Modelo.PropArea.GetNameList(ref number, ref names);

        //    eSlabType slabType = 0;
        //    eShellType eShell = 0;

        //    for (int i = 0; i < names.Length; i++)
        //    {
        //        string MathProp = ""; double Thinkness = 0; int color = 0; string note = ""; string GuID = "";
        //        Modelo.PropArea.GetSlab(names[i], ref slabType, ref eShell, ref MathProp, ref Thinkness, ref color, ref note, ref GuID);

        //        if (GuID != "")
        //        {
        //            NamesSlabs.Add(names[i]);
        //        }
        //    }
        //}

        public static void Set_Load_Cases(List<string> Load_Combos, ref cAnalysisResults Resultados, int tipo)
        {
            int Numero_Elementos = 0;
            string[] Casos_Carga = null;
            string prueba;

            Resultados.Setup.DeselectAllCasesAndCombosForOutput();

            if (tipo == 0)
            {
                Modelo.LoadCases.GetNameList(ref Numero_Elementos, ref Casos_Carga);
            }
            else
            {
                Modelo.RespCombo.GetNameList(ref Numero_Elementos, ref Casos_Carga);
            }

            for (int i = 0; i < Load_Combos.Count; i++)
            {
                prueba = Casos_Carga.ToList().Find(x => x.ToUpper() == Load_Combos[i].ToUpper());
                if (tipo == 0)
                {
                    Resultados.Setup.SetCaseSelectedForOutput(prueba, true);
                }
                else
                {
                    Resultados.Setup.SetComboSelectedForOutput(prueba, true);
                }
            }
            //Resultados.Setup.SetOptionMultiStepStatic(2);
            //Resultados.Setup.SetOptionMultiValuedCombo(2);
        }

        //public static void Add_Load_Case(string Loadcase)
        //{
        //    cLoadPatterns CasosCarga;
        //    int Num_elem = 0;
        //    var Load_cases = new string[] { };

        //    CasosCarga = Modelo.LoadPatterns;
        //    CasosCarga.GetNameList(ref Num_elem, ref Load_cases);

        //    if (Load_cases.ToList().Exists(x => x == Loadcase) == false) CasosCarga.Add(Loadcase, eLoadPatternType.Live, 0);
        //}

        //public static void Set_area_loads(string Piso, double Factor)
        //{
        //    int Num_elem = 0;
        //    var Areas_piso = new string[] { };
        //    var Area_names = new string[] { };
        //    var Load_cases = new string[] { };
        //    var CSys = new string[] { };
        //    var dir = new int[] { };
        //    var Carga = new double[] { };
        //    var Design_orientation = new eAreaDesignOrientation();
        //    double Carga_O;

        //    cAreaObj Areas = Modelo.AreaObj;
        //    Areas.GetNameListOnStory(Piso, ref Num_elem, ref Areas_piso);

        //    foreach (string Area_i in Areas_piso)
        //    {
        //        Areas.GetDesignOrientation(Area_i, ref Design_orientation);

        //        if (Design_orientation == eAreaDesignOrientation.Floor)
        //        {
        //            Areas.GetLoadUniform(Area_i, ref Num_elem, ref Area_names, ref Load_cases, ref CSys, ref dir, ref Carga);

        //            if (Load_cases.ToList().Exists(x => x == "LIVE_O") == false)
        //            {
        //                for (int i = 0; i < Load_cases.Count(); i++)
        //                {
        //                    if (Load_cases[i] == "LIVE")
        //                    {
        //                        Carga_O = Carga[i];
        //                        Areas.DeleteLoadUniform(Area_i, Load_cases[i]);
        //                        Areas.SetLoadUniform(Area_i, Load_cases[i], Carga_O * Factor, dir[i]);
        //                        Areas.SetLoadUniform(Area_i, "LIVE_O", Carga_O, dir[i]);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

        //public static void Set_frame_loads(string Piso, double Factor)
        //{
        //    int Num_elem = 0;
        //    var Frames_piso = new string[] { };
        //    var Frames_names = new string[] { };
        //    var Load_cases = new string[] { };
        //    var Tipo_carga = new int[] { };
        //    var CSys = new string[] { };
        //    var dir = new int[] { };
        //    var Carga = new double[] { };
        //    var Csys = new string[] { };
        //    var FrameName = new string[] { };
        //    var Design_orientation = new eFrameDesignOrientation();
        //    var RD1 = new double[] { };
        //    var RD2 = new double[] { };
        //    var Dist1 = new double[] { };
        //    var Dist2 = new double[] { };
        //    var Val1 = new double[] { };
        //    var Val2 = new double[] { };

        //    double Val_1, Val_2;

        //    cFrameObj Frames = Modelo.FrameObj;
        //    cAnalysisResults Analisis = Modelo.Results;
        //    Frames.GetNameListOnStory(Piso, ref Num_elem, ref Frames_piso);

        //    Set_Load_Cases(new string[] { "DEAD", "LIVE" }.ToList(), ref Analisis, 0);

        //    foreach (string Framei in Frames_piso)
        //    {
        //        Frames.GetDesignOrientation(Framei, ref Design_orientation);
        //        if (Design_orientation == eFrameDesignOrientation.Beam)
        //        {
        //            Frames.GetLoadDistributed(Framei, ref Num_elem, ref FrameName, ref Load_cases, ref Tipo_carga, ref CSys, ref dir, ref RD1, ref RD2, ref Dist1, ref Dist2, ref Val1, ref Val2);

        //            if (Load_cases.ToList().Exists(x => x == "LIVE_O") == false)
        //            {
        //                for (int i = 0; i < Load_cases.Count(); i++)
        //                {
        //                    if (Load_cases[i] == "LIVE")
        //                    {
        //                        Val_1 = Val1[i]; Val_2 = Val2[i];
        //                        Frames.DeleteLoadDistributed(Framei, Load_cases[i]);
        //                        Frames.SetLoadDistributed(Framei, Load_cases[i], 1, 6, 0, 1, Val_1 * Factor, Val_2 * Factor, "Global", true, true, 0);
        //                        Frames.SetLoadDistributed(Framei, "LIVE_O", 1, 6, 0, 1, Val_1, Val_2, "Global", true, true, 0);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

        //public static void Get_area(Listas_Objetos Lista_i, string Piso)
        //{
        //    int NumberNames = 0;
        //    string[] MyName = null;
        //    string Pier = "None";
        //    string propiedad = "";
        //    string[] Puntos = null;
        //    double X = 0;
        //    double Y = 0;
        //    double Z = 0;
        //    int indice = 0;

        //    Pier_Shell Shell_i;
        //    cAreaObj Areas;

        //    Areas = Modelo.AreaObj;
        //    Areas.GetNameList(ref NumberNames, ref MyName);

        //    foreach (string Area_i in MyName)
        //    {
        //        Areas.GetPier(Area_i, ref Pier);
        //        if (Pier != "None")
        //        {
        //            Areas.GetProperty(Area_i, ref propiedad);
        //            Shell_i = new Pier_Shell
        //            {
        //                Label = Area_i,
        //                Material = propiedad,
        //                Pier = Pier
        //            };

        //            Shell_i.Coordenadas = new List<double[]>();
        //            Areas.GetPoints(Area_i, ref NumberNames, ref Puntos);

        //            foreach (string punto_i in Puntos)
        //            {
        //                Modelo.PointObj.GetCoordCartesian(punto_i, ref X, ref Y, ref Z);
        //                Shell_i.Coordenadas.Add(new double[] { X, Y, Z });
        //            }

        //            indice = Lista_i.Lista_Muros.FindIndex(x => x.Label == Shell_i.Pier & x.Story == Piso);
        //            Lista_i.Lista_Muros[indice].Shells.Add(Shell_i);
        //        }
        //    }
        //    for (int i = 0; i < Lista_i.Lista_Muros.Count; i++)
        //    {
        //        Lista_i.Lista_Muros[i].Get_Extremos();
        //    }
        //}

        public static void Get_ModalPeriods(ref double[] Periods, ref double[] UX, ref double[] UY, ref double[] UZ, ref double[] RX, ref double[] RY, ref double[] RZ)
        {
            Application.UseWaitCursor = true;
            if (Modelo.GetModelIsLocked() == false) Modelo.Analyze.RunAnalysis();

            int NumberResults = 0;
            string[] LoadCases = null; string[] StepTep = null; double[] stepNum = null;
            double[] sumUX = null; double[] sumUY = null; double[] sumUZ = null;
            double[] sumRX = null; double[] sumRY = null; double[] sumRZ = null;

            List<string> Load_Cases = new List<string>();
            cAnalysisResults Analisis = Modelo.Results;

            Modelo.Results.Setup.DeselectAllCasesAndCombosForOutput();


            Load_Cases.AddRange(new string[] { "Modal" });

            Set_Load_Cases(Load_Cases, ref Analisis, 0);
            Analisis.ModalParticipatingMassRatios(ref NumberResults, ref LoadCases, ref StepTep, ref stepNum, ref Periods, ref UX, ref UY, ref UZ, ref sumUX, ref sumUY, ref sumUZ, ref RX, ref RY, ref RZ, ref sumRX, ref sumRY, ref sumRZ);

            Application.UseWaitCursor = false;
        }


        public static void ObtenerAluturaEdificio(ref string[] StoryNames, ref double[] StoryElevations, ref double[] StoryHeights)
        {
            int NumeroDePisos = 0;
            bool[] IsMasterStory = { };
            string[] SimilarToStory = { };
            bool[] SpliceAbove = { };
            double[] SpliceHeight = { };

            Modelo.Story.GetStories(ref NumeroDePisos, ref StoryNames, ref StoryElevations, ref StoryHeights, ref IsMasterStory, ref SimilarToStory, ref SpliceAbove, ref SpliceHeight);

        }

        //public static void Serializar(Listas_Objetos Lista_i)
        //{
        //    Serializador.Serializar(Lista_i);
        //}

        //public static List<string> Secciones()
        //{
        //    int num_items = 0;
        //    string[] FramesNames = { };
        //    cPropFrame propFrame = Modelo.PropFrame;
        //    List<string> Temp = new List<string>();

        //    propFrame.GetNameList(ref num_items, ref FramesNames);
        //    Temp = FramesNames.ToList();

        //    return Temp;
        //}

        //public static void Asignar_frame(string frameLabel, string FrameSection)
        //{
        //    if (Modelo.GetModelIsLocked() == true) Modelo.SetModelIsLocked(false);
        //    Modelo.FrameObj.SetSection(frameLabel, FrameSection);
        //}

        //public static void Get_Type_Frames(eFrameDesignOrientation designOrientation, ref List<Tuple<string, string, string, double[], double>> Frame_prop)
        //{
        //    int NumberNames = 0;
        //    string PropName = "";
        //    int Cont = 0;
        //    string SAuto = "";
        //    string[] MyName = { };
        //    string[] MyLabel = { };
        //    string[] MyStory = { };
        //    double[] Aux_Coord = new double[3];
        //    double Xc = 0; double Yc = 0; double Zc = 0;
        //    string P1 = ""; string P2 = "";
        //    eFrameDesignOrientation FrameDesign = eFrameDesignOrientation.Null;

        //    #region Frame_properties

        //    double Area = 0;
        //    double As2 = 0;
        //    double As3 = 0;
        //    double Torsion = 0;
        //    double I22 = 0;
        //    double I33 = 0;
        //    double S22 = 0;
        //    double S33 = 0;
        //    double Z22 = 0;
        //    double Z33 = 0;
        //    double R22 = 0;
        //    double R33 = 0;

        //    #endregion Frame_properties

        //    Modelo.FrameObj.GetLabelNameList(ref NumberNames, ref MyName, ref MyLabel, ref MyStory);

        //    foreach (string Framei in MyName)
        //    {
        //        Modelo.FrameObj.GetDesignOrientation(Framei, ref FrameDesign);

        //        if (FrameDesign == designOrientation)
        //        {
        //            Modelo.FrameObj.GetSection(Framei, ref PropName, ref SAuto);
        //            Modelo.PropFrame.GetSectProps(PropName, ref Area, ref As2, ref As3, ref Torsion, ref I22, ref I33, ref S22, ref S33, ref Z22, ref Z33, ref R22, ref R33);
        //            Modelo.FrameObj.GetPoints(Framei, ref P1, ref P2);
        //            Modelo.PointObj.GetCoordCartesian(P1, ref Xc, ref Yc, ref Zc);
        //            Aux_Coord = new double[] { Math.Round(Xc, 2), Math.Round(Yc, 2), Math.Round(Zc, 2) };
        //            Frame_prop.Add(new Tuple<string, string, string, double[], double>(Framei, MyLabel[Cont], PropName, Aux_Coord, Area));
        //        }
        //        Cont++;
        //    }
        //}

        //public static void Get_Frame_Design(eFrameDesignOrientation designOrientation, string Frame_name, ref double Area_Ref)
        //{
        //    #region Variables_Columnas

        //    int NumberItems = 0;
        //    string[] FrameName = { };
        //    int[] MyOption = { };
        //    double[] Location = { };
        //    string[] PMMCombo = { };
        //    double[] PMMArea = { };
        //    double[] PMMRatio = { };
        //    string[] VMajorCombo = { };
        //    double[] AVMajor = { };
        //    string[] VMinorCombo = { };
        //    double[] AVMinor = { };
        //    string[] ErrorSummary = { };
        //    string[] WarningSummary = { };
        //    eItemType ItemType = eItemType.Objects;

        //    #endregion Variables_Columnas

        //    #region Variables_Vigagas

        //    string[] TopCombo = { };
        //    double[] TopArea = { };
        //    string[] BotCombo = { };
        //    double[] BotArea = { };
        //    double[] VMajorArea = { };
        //    string[] TLCombo = { };
        //    double[] TLArea = { };
        //    string[] TTCombo = { };
        //    double[] TTArea = { };

        //    #endregion Variables_Vigagas

        //    #region Variables Seccion

        //    string PropName = "";
        //    string SAuto = "";
        //    string MatPropLong = "";
        //    string MatPropConfine = "";
        //    int Pattern = 0;
        //    int ConfineType = 0;
        //    double Cover = 0.0;
        //    int NumberCBars = 0;
        //    int NumberR3Bars = 0;
        //    int NumberR2Bars = 0;
        //    string RebarSize = "";
        //    string TieSize = "";
        //    double TieSpacingLongit = 0.0;
        //    int Number2DirTieBars = 0;
        //    int Number3DirTieBars = 0;
        //    bool ToBeDesigned = true;
        //    string LongitCornerRebarSize = "";
        //    double LongitRebarArea = 0.00;
        //    double LongitCornerRebarArea = 0.00;

        //    #endregion Variables Seccion

        //    if (Modelo.DesignConcrete.GetResultsAvailable() == false)
        //    {
        //        if (Modelo.GetModelIsLocked() == false) Modelo.Analyze.RunAnalysis();
        //        Modelo.DesignConcrete.StartDesign();
        //    }

        //    if (designOrientation == eFrameDesignOrientation.Column)
        //    {
        //        Modelo.DesignConcrete.GetSummaryResultsColumn(Frame_name, ref NumberItems, ref FrameName, ref MyOption, ref Location,
        //            ref PMMCombo, ref PMMArea, ref PMMRatio, ref VMajorCombo, ref AVMajor, ref VMinorCombo, ref AVMinor, ref ErrorSummary,
        //            ref WarningSummary, ItemType);

        //        if (PMMArea.ToList().Exists(x => x > 0))
        //        {
        //            Area_Ref = PMMArea.Max() * Math.Pow(100f, 2);
        //        }
        //        else
        //        {
        //            Modelo.PropFrame.GetRebarColumn(PropName, ref MatPropLong, ref MatPropConfine, ref Pattern, ref ConfineType,
        //                ref Cover, ref NumberCBars, ref NumberR3Bars, ref NumberR2Bars, ref RebarSize, ref TieSize, ref TieSpacingLongit,
        //                ref Number2DirTieBars, ref Number3DirTieBars, ref ToBeDesigned);
        //            Modelo.FrameObj.GetSection(Frame_name, ref PropName, ref SAuto);

        //            Modelo.PropFrame.GetRebarColumn_1(PropName, ref MatPropLong, ref MatPropConfine, ref Pattern, ref ConfineType,
        //              ref Cover, ref NumberCBars, ref NumberR3Bars, ref NumberR2Bars, ref RebarSize, ref TieSize, ref TieSpacingLongit,
        //             ref Number2DirTieBars, ref Number3DirTieBars, ref ToBeDesigned, ref LongitCornerRebarSize, ref LongitRebarArea, ref LongitCornerRebarArea);
        //            Area_Ref = NumberCBars * LongitRebarArea * Math.Pow(100f, 2);
        //        }
        //    }

        //    if (designOrientation == eFrameDesignOrientation.Beam)
        //    {
        //        Modelo.DesignConcrete.GetSummaryResultsBeam(Frame_name, ref NumberItems, ref FrameName, ref Location, ref TopCombo, ref TopArea, ref BotCombo, ref BotArea, ref VMajorCombo,
        //            ref VMajorArea, ref TLCombo, ref TLArea, ref TTCombo, ref TTArea, ref ErrorSummary, ref WarningSummary);
        //    }
        //}
    }
}

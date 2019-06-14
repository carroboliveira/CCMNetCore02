//==============================================
//Classe     : ClassStarLine
//Descrição  : Metodos para ser utilizados na automação de teste
//Criado por : Carlos R. S. Oliveira
//Criado em  : 20/04/2019
//==============================================
using Npgsql;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace Starline
{
    public class ProcessTest
    {
        //public string wpath = PastaBase "C:\\ProgramData\\SISTEMAS\\Prints\\PortalExtranet";
        static ConexaoNpgSql Conexao;
        public int CstID { get; set; }
        public int SitID { get; set; }
        public int ScnID { get; set; }
        public int TstID { get; set; }
        public int RptID { get; set; }
        public int ReportID { get; set; }
        public int LgsID { get; set; }
        public int StepNumber { get; set; }
        public int StepTurn { get; set; }
        public string StepName { get; set; }
        public string CustomerName { get; set; }
        public string SuiteName { get; set; }
        public string ScenarioName { get; set; }
        public string TestName { get; set; }
        public string TestType { get; set; }
        public string AnalystName { get; set; }
        public string TestDesc { get; set; }
        public string PreCondition { get; set; }
        public string PostCondition { get; set; }
        public string InputData { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime FinishTime { get; set; }

        public string PastaBase { get; set; }

        public bool TemConexao = false;
        public Dictionary<string, Dictionary<string, int>> ListTurnos;



        public ProcessTest()
        {
            Conexao = new ConexaoNpgSql();


            try
            {
                Conexao.AbrirConexao();
                TemConexao = Conexao.ConexaoAberta();
                Conexao.FecharConexao();
            }
            catch
            {
                TemConexao = false;
            }
            ListTurnos = new Dictionary<string, Dictionary<string, int>>();
        }



        private string Print(string NomeMetodo, Exception ex)
        {
            string msg;
            if (ex != null)
            {
                msg = NomeMetodo + " ( " + ex + " )";

            }
            else
            {
                msg = NomeMetodo;
            }
            Console.Error.WriteLine(msg);
            return msg;
        }

        #region LisTurnos
        public int NovoTurno(string TurnoOne, string TurnoTwo)
        {
            try
            {
                ListTurnos[TurnoOne][TurnoTwo] += 1;
                return ListTurnos[TurnoOne][TurnoTwo];
            }
            catch
            {
                if (ListTurnos.ContainsKey(TurnoOne))
                {
                    ListTurnos[TurnoOne].Add(TurnoTwo, 1);
                }
                else
                {
                    ListTurnos.Add(TurnoOne, new Dictionary<string, int>() { { TurnoTwo, 1 } });
                }
                return ListTurnos[TurnoOne][TurnoTwo];
            }
        }


        #endregion ListTurnos

        //Descontinuado
        //private int Dml(string sql, string Returning = "")
        //{
        //    int id = 0;
        //    try
        //    {
        //        if (Returning.Length > 0)
        //        {
        //            string pSql = sql;
        //            string tabela = pSql.Substring(pSql.IndexOf("from") + 4, pSql.IndexOf("where") - 5);
        //            string condicao = pSql.Remove(0, pSql.IndexOf("where"));
        //            string select = "Select " + Returning + " from " + tabela + " where " + condicao;
        //            NpgsqlDataReader dr = Conexao.ExecutarSelect(select);
        //            if (dr.HasRows && dr.Read())
        //            {
        //                id = Convert.ToInt32(dr[Returning].ToString());
        //            }
        //        }

        //        Conexao.ExecutarScript(sql);
        //{
        //        if (id != 0)
        //        {
        //            return id;
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(Print("Dml", ex));
        //    }
        //    return 0;
        //}

        //Descontinuado
        //private NpgsqlDataReader Query(string sql)
        //{
        //    try
        //    {
        //        return Conexao.ExecutarSelect(sql);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(Print("Query", ex));
        //    }

        //}

        public string RemoveAcentos(string txt)
        {
            if (string.IsNullOrEmpty(txt))
            {
                return string.Empty;
            }
            Byte[] bytes = Encoding.GetEncoding("iso-8859-8").GetBytes(txt);
            return Encoding.UTF8.GetString(bytes);
        }
        public string Slugify(string value)
        {
            try
            {
                //Passar para letras minusculas 
                value = value.ToLowerInvariant();

                //Tiro todos os acentos
                value = RemoveAcentos(value);

                //Substituo todos os espaços
                value = Regex.Replace(value, @"\s", "-", RegexOptions.Compiled);

                //Removo alguns caracteres especiais, se houver
                value = Regex.Replace(value, @"[^\w\s\p{Pd}]", "", RegexOptions.Compiled);

                //Tiro os Traços do final  
                value = value.Trim('-', '_');

                //Tiro ocorrencias duplas de - ou \_ 
                value = Regex.Replace(value, @"([-_]){2,}", "$1", RegexOptions.Compiled);

                //e retorno o novo valor
                return value;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Slugify", ex));
            }
        }

        public string Call(string NomeDaFuncao, params string[] ListaParam)
        {
            if (TemConexao)
            {
                Conexao.AbrirConexao();
                if (Conexao.ConexaoAberta())
                {
                    try
                    {
                        string select = "select " + NomeDaFuncao + "(";
                        for (var x = 0; x <= ListaParam.Length - 1; x++)
                        {

                            if (x == 0)
                            {
                                select += "@CP" + x.ToString();
                            }
                            else
                            {
                                select += ",@CP" + x.ToString();
                            }
                        }
                        select += ") as ID";

                        NpgsqlCommand cmd = new NpgsqlCommand(select, Conexao.Conn);
                        for (var x = 0; x <= ListaParam.Length - 1; x++)
                        {

                            string Conteudo = ListaParam[x].Substring(0, ListaParam[x].IndexOf(":"));
                            string Value = ListaParam[x].Remove(0, ListaParam[x].IndexOf(":") + 1);
                            if (Conteudo.ToLower().IndexOf("_id") > 0)
                            {
                                cmd.Parameters.AddWithValue("CP" + x.ToString(), Convert.ToInt32(Value));
                            }
                            else
                            {
                                cmd.Parameters.AddWithValue("CP" + x.ToString(), Value);
                            }
                        }
                        string Result = "";
                        NpgsqlDataReader dr = cmd.ExecuteReader();
                        if (dr.HasRows && dr.Read())
                        {
                            Result = dr["ID"].ToString();
                        }
                        dr.Close();
                        return Result;

                    }
                    catch (Exception ex)
                    {
                        throw new Exception(Print("Call", ex));
                    }

                }
                else
                {
                    return "1";
                }
            }
            else
            {
                return "1";
            }

        }


        //Descontinuada
        //private int Convert.ToInt32(NpgsqlDataReader dr)
        //{
        //    if (dr != null)
        //    {
        //        if (dr.HasRows && dr.Read())
        //        {
        //            int id = Convert.ToInt32(dr["ID"].ToString());
        //            dr.Close();
        //            return id;
        //        }
        //        else
        //        {
        //            return 0;
        //        }

        //    }
        //    else
        //    {
        //        return 0;
        //    }
        //}
        public int StartTest(string customerName, string suiteName, string scenarioName, string testName, string testType, string analystName, string testDesc, int ReportID = 0, int deleteFlag = 1)
        {
            Print("", null);
            customerName = Slugify(customerName);
            CstID = Convert.ToInt32(Call("arch.update_cst", "customer:" + customerName));
            Print("Customer ID: " + CstID.ToString() + " [ " + customerName + " ]", null);
            SuiteName = Slugify(suiteName);
            SitID = Convert.ToInt32(Call("auto.update_sit", "cst_id:" + CstID.ToString(), "suite:" + SuiteName));
            Print("Suite ID: " + SitID.ToString() + " [ " + suiteName + " ]", null);
            ScenarioName = Slugify(scenarioName);
            ScnID = Convert.ToInt32(Call("auto.update_scn", "cst_id:" + CstID.ToString(), "scenario:" + ScenarioName));
            Print("Scenario ID: " + ScnID.ToString() + " [ " + ScenarioName + " ]", null);
            TestName = Slugify(testName);
            TestType = testType;
            TestDesc = testDesc;
            AnalystName = analystName;
            TstID = Convert.ToInt32(Call("auto.update_tst", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "test:" + TestName, "type:" + TestType, "desc:" + TestDesc, "analyst:" + AnalystName));
            Print("Test ID: " + TstID.ToString() + " [ " + TestName + " ]", null);
            Print("", null);

            if (ReportID != 0 && deleteFlag == 1)
            {
                Call("auto.delete_rpt", "cst_id:" + CstID.ToString(), "rpt_id:" + ReportID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString());
            }
            if (TemConexao)
            {
                RptID = Convert.ToInt32(Call("auto.update_rpt", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "rpt_id:" + ReportID.ToString(), "status:pendente"));
                if (RptID == 0)
                {
                    return 0;
                };
            }
            else
            {
                RptID = ReportID;
            }
            Print("Report ID: " + RptID.ToString(), null);
            if (deleteFlag == 1)
            {
                Call("auto.delete_lgs", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "rpt_id:" + RptID.ToString());
                Call("auto.delete_stp", "cst_id:" + CstID.ToString(), "tst_id:" + TstID.ToString());
            }
            return RptID;

        }
        public void DoTest(string pre = "", string post = "", string inputData = "", int steps = 0)
        {
            try
            {
                if (this.CstID == 0)
                {
                    throw new Exception(Print("DoTest [ CstID=0 ]", null));
                }
                if (this.SitID == 0)
                {
                    throw new Exception(Print("DoTest [ SitID=0 ]", null));
                }
                if (this.ScnID == 0)
                {
                    throw new Exception(Print("DoTest [ ScnID=0 ]", null));
                }
                if (this.TstID == 0)
                {
                    throw new Exception(Print("DoTest [ TstID=0 ]", null));
                }
                Call("auto.update_tst_details", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id: " + TstID.ToString(), "pre:" + pre, "post:" + post, "input_data:" + inputData, "steps_id:" + steps.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception(Print("DoTest", ex));
            }

        }
        public int DoStep(string desc, string expected_result = "", string status = "active", int newStep = 1)
        {
            try
            {
                if (this.CstID == 0)
                {
                    throw new Exception(Print("DoStep(CstID=0)", null));
                }
                if (this.SitID == 0)
                {
                    throw new Exception(Print("DoStep(SitID=0)", null));
                }
                if (this.ScnID == 0)
                {
                    throw new Exception(Print("DoStep(ScnID=0)", null));
                }
                if (this.TstID == 0)
                {
                    throw new Exception(Print("DoStep(TstID=0)", null));
                }
                if (this.RptID == 0)
                {
                    throw new Exception(Print("DoStep(RptID=0)", null));
                }
                StepNumber = Convert.ToInt32(Call("auto.update_stp", "cst_id:" + CstID.ToString(), "tst_id:" + TstID.ToString(), "dsc:" + desc, "expected_result:" + expected_result, "status:" + status));
                if (newStep == 1)
                {
                    Call("auto.delete_lgs", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "rpt_id:" + RptID.ToString(), "description:" + desc);
                }
                Call("auto.count_stp", "cst_id:" + CstID.ToString(), "tst_id:" + TstID.ToString());
                StartStep(desc, 1, status = "pendente");
                return StepNumber;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("DoStep", ex));
            }

        }
        public int StartStep(string desc, int turn = 1, string status = "executando", string logMsg = "", string paramName = "", string paramValue = "")
        {
            try
            {
                if (this.CstID == 0)
                {
                    throw new Exception(Print("StartStep(CstID=0)", null));
                }
                if (this.SitID == 0)
                {
                    throw new Exception(Print("StartStep(SitID=0)", null));
                }
                if (this.ScnID == 0)
                {
                    throw new Exception(Print("StartStep(ScnID=0)", null));
                }
                if (this.TstID == 0)
                {
                    throw new Exception(Print("StartStep(TstID=0)", null));
                }
                if (this.RptID == 0)
                {
                    throw new Exception(Print("StartStep(RptID=0)", null));
                }

                if (status == "executando")
                {

                    string logLine = "  Step " + GetStepNumber(desc).ToString() + "." + turn.ToString() + ": " + desc;
                    if (logMsg != "" || paramName != "")
                    {
                        logLine = logLine + " (";
                        if (logMsg != "")
                        {
                            logLine = logLine + "Obs: " + logMsg;
                            if (paramName != "")
                            {
                                logLine = logLine + " - ";
                            }
                        }
                        if (paramName != "")
                        {
                            logLine = logLine + "Parameter: " + paramName + " = " + paramValue;
                            logLine = logLine + ")";
                        }

                    }
                    Print("StartStep('')", null);
                    Print("StartStep(" + logLine + ")", null);
                }
                return Convert.ToInt32(Call("auto.update_lgs", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "rpt_id:" + RptID.ToString(), "dsc:" + desc, "turn_id:" + turn.ToString(), "status:" + status, "log_msg:" + logMsg, "param_name:" + paramName, "param_value:" + paramValue));
            }
            catch (Exception ex)
            {
                throw new Exception(Print("StartStep", ex));
            }

        }
        public int GetStepNumber(string stepName)
        {
            try
            {
                if (this.CstID == 0)
                {
                    throw new Exception(Print("GetStepNumber(CstID=0)", null));
                }
                if (this.SitID == 0)
                {
                    throw new Exception(Print("GetStepNumber(SitID=0)", null));
                }
                return Convert.ToInt32(Call("auto.get_stp_num", "cst_id:" + CstID.ToString(), "tst_id:" + TstID.ToString(), "desc:" + stepName));
            }
            catch (Exception ex)
            {
                throw new Exception(Print("GetStepNumber", ex));
            }

        }
        public bool EndStep(int lgsID, string status = "sucesso", string PrintPath = "", string logMsg = "", string reason = "")
        {
            try
            {
                if (this.CstID == 0)
                {
                    throw new Exception(Print("EndStep[CstID=0]", null));
                }
                if (this.RptID == 0)
                {
                    throw new Exception(Print("EndStep[RptID=0]", null));
                }

                if (logMsg != "")
                {
                    if (status == "sucesso")
                    {
                        Print("    EndStep(Obs: " + logMsg + ")", null);
                    }
                    else
                    {
                        Print("    EndStep(" + logMsg + ")", null);

                    }

                }
                PrintPath = PrintPath.Replace(@"\", @"/");
                Call("auto.update_lgs_details", "cst_id:" + CstID.ToString(), "lgs_id:" + lgsID.ToString(), "status:" + status, "Print_path:" + PrintPath, "log_msg:" + logMsg, "reason:" + reason);
                PutImage(CstID, RptID, lgsID, PrintPath);
                if (status == "erro" || status == "grave")
                {
                    return false;
                }
                else
                {
                    return true;
                }

            }
            catch (Exception ex)
            {
                return false;
                throw new Exception(Print("EndStep", ex));
            }

        }

        public string PrintPage(int sleep = 0)
        {
            if (sleep > 0)
            {
                System.Threading.Thread.Sleep(sleep * 1000);
            }
            //Random random = new Random();
            //int randID = random.Next(1, 1000000);

            string wpath = "C:\\ProgramData\\SISTEMAS\\Prints\\PortalExtranet";
            //Se o diretório não existir...
            if (!Directory.Exists(wpath))
            {
                //Criamos um com o nome folder
                Directory.CreateDirectory(wpath);
            }

            string _step = this.StepNumber.ToString().PadLeft(4, '0');
            string _seq = this.StepTurn.ToString().PadLeft(2, '0');
            string FileName = this.TestName.Replace(" ", "_");
            string filename = wpath + "\\" + _step + "_" + _seq + "-" + FileName + ".png";


            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(printscreen as Image);
            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);
            printscreen.Save(filename, ImageFormat.Png);
            return filename;
        }


        public bool PutImage(int CstId, int RptId, int lgsID, string file_full_path)
        {
            try
            {
                if (TemConexao)
                {
                    Conexao.AbrirConexao();
                    if (Conexao.ConexaoAberta())
                    {
                        NpgsqlCommand cmd = new NpgsqlCommand("select auto.put_image(@CstId,@RptId,@LogsId,@ImageData) as ID", Conexao.Conn);
                        cmd.Parameters.AddWithValue("ImageData", ImageToByte(Image.FromFile(file_full_path)));
                        cmd.Parameters.AddWithValue("LogsId", lgsID);
                        cmd.Parameters.AddWithValue("CstId", CstId);
                        cmd.Parameters.AddWithValue("RptId", RptId);
                        cmd.ExecuteNonQuery();

                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Print("PutImage", ex);
                return false;
            }
        }
        public bool PutImage_old(int CstId, int RptId, int lgsID, string file_full_path)
        {
            try
            {
                Call("auto.put_image", "cst_id:" + CstID.ToString(), "rpt_id:" + RptId.ToString(), "lgs_id:" + lgsID.ToString(), "image_data:" + ImageToByte(Image.FromFile(file_full_path)));
                return true;
            }
            catch (Exception ex)
            {
                Print("PutImage", ex);
                return false;
            }
        }

        //Descontinuado
        //public void GetImage(int CstId, int RptId, int lgsID, string file_full_path)
        //{
        //    NpgsqlDataReader dr = null;
        //    try
        //    {
        //        //NpgsqlDataReader dr = Query("select image_data from auto.images where cst_id = "+CstId.ToString()+" and rpt_id =  "+RptId.ToString()+" and lgs_id = " + lgsID.ToString());
        //        dr = Call("auto.get_image", "cst_id:" + CstId.ToString(), "rpt_id:" + RptId.ToString(), "lgs_id:" + lgsID.ToString());
        //        if (dr.HasRows && dr.Read())
        //        {
        //            Base64ToImage(dr["image_data"].ToString(), file_full_path);
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        Print("GetImage", ex);
        //    }
        //    finally
        //    {
        //        dr.Close();
        //    }

        //}
        public string GetImageBase64(string PathImage)
        {
            try
            {
                Image img = Image.FromFile(PathImage);
                MemoryStream ms = new MemoryStream();
                img.Save(ms, img.RawFormat);
                byte[] ImgEmBytes = ms.ToArray();
                return Convert.ToBase64String(ImgEmBytes);
            }
            catch (Exception ex)
            {
                Print("GetImageBase64", ex);
                return "";
            }
        }
        public bool Base64ToImage(string ImagemBase64, string file_full_path)
        {
            try
            {
                byte[] imgBytes = Convert.FromBase64String(ImagemBase64);
                MemoryStream ms = new MemoryStream(imgBytes, 0, imgBytes.Length);
                ms.Write(imgBytes, 0, imgBytes.Length);
                Image img = Image.FromStream(ms, true);
                img.Save(file_full_path);
                return true;

            }
            catch (Exception ex)
            {
                Print("Base64ToImage", ex);
                return false;
            }
        }
        public byte[] ImageToByte(Image PrintScreen)
        {
            if (PrintScreen != null)
            {
                MemoryStream ms = new MemoryStream();
                PrintScreen.Save(ms, ImageFormat.Png);
                ms.Seek(0, SeekOrigin.Begin);
                byte[] imgByte = new byte[ms.Length];
                ms.Read(imgByte, 0, Convert.ToInt32(ms.Length));
                return imgByte;
            }
            return null;
        }

        //Descontinuado
        //public Image ByteToImage(int CstId, int RptId, int lgsID, string file_full_path)
        //{
        //    try
        //    {

        //        NpgsqlDataReader dr = Call("auto.get_image", "cst_id:" + CstId.ToString(), "rpt_id:" + RptId.ToString(), "lgs_id:" + lgsID.ToString());
        //        if (dr.HasRows && dr.Read())
        //        {
        //            MemoryStream ms = new MemoryStream((byte[])dr["image_data"]);
        //            Image img = Image.FromStream(ms, true);
        //            dr.Close();
        //            return img;
        //        }
        //        else
        //        {
        //            dr.Close();
        //            return null;
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        Print("ByteToImage", ex);
        //        return null;
        //    }
        //}

        public void EndTest(int ReportID = 0, string status = "pronto", bool deleteReport = false)
        {
            try
            {
                if (this.SitID == 0)
                {
                    throw new Exception(Print("EndTest(SitID=0)", null));
                }
                if (this.ScnID == 0)
                {
                    throw new Exception(Print("EndTest(ScnID=0)", null));
                }
                if (this.TstID == 0)
                {
                    throw new Exception(Print("EndTest(TstID=0)", null));
                }
                if (this.RptID == 0)
                {
                    throw new Exception(Print("EndTest(RptID=0)", null));
                }

                if (deleteReport == true || ReportID == 0)
                {
                    Call("auto.delete_rpt", "cst_id:" + CstID.ToString(), "rpt_id:" + RptID.ToString());
                }

                if (Conexao.ConexaoAberta())
                {
                    Conexao.FecharConexao();
                    Print("End Test", null);
                }

            }
            catch (Exception ex)
            {
                Print("EndTest", ex);
            }
        }

        public string PrintPageComSelenium(IWebDriver DriverDoSelenium, bool Full = false, int sleep = 0)
        {
            try
            {
                string wpath = this.PastaBase;

                if (this.CustomerName.Length > 0)
                {
                    wpath += "/" + Slugify(this.CustomerName);
                }
                if (this.ReportID.ToString().Length > 0)
                {
                    wpath += "/" + Slugify(this.ReportID.ToString());
                }
                if (this.SuiteName.Length > 0)
                {
                    wpath += "/" + Slugify(this.SuiteName);
                }
                if (this.ScenarioName.Length > 0)
                {
                    wpath += "/" + Slugify(this.ScenarioName);
                }
                if (this.TestName.Length > 0)
                {
                    wpath += "/" + Slugify(this.TestName);
                }

                if (sleep > 0)
                {
                    System.Threading.Thread.Sleep(sleep * 1000);
                }


                //Se o diretório não existir...
                if (!Directory.Exists(wpath))
                {
                    //Criamos um com o nome folder
                    Directory.CreateDirectory(wpath);
                }

                string _step = this.StepNumber.ToString().PadLeft(4, '0');
                string _seq = this.StepTurn.ToString().PadLeft(2, '0');
                string FileName = this.TestName.Replace(" ", "_");
                string filename = wpath + "\\" + _step + "_" + _seq + "-" + FileName + ".png";


                if (File.Exists(filename))
                {
                    File.Delete(filename);
                }

                if (Full)
                {
                    Bitmap stitchedImage = null;
                    long totalwidth1 = (long)((IJavaScriptExecutor)DriverDoSelenium).ExecuteScript("return document.body.offsetWidth");//documentElement.scrollWidth");

                    long totalHeight1 = (long)((IJavaScriptExecutor)DriverDoSelenium).ExecuteScript("return  document.body.parentNode.scrollHeight");

                    int totalWidth = (int)totalwidth1;
                    int totalHeight = (int)totalHeight1;

                    // Get the Size of the Viewport
                    long viewportWidth1 = (long)((IJavaScriptExecutor)DriverDoSelenium).ExecuteScript("return document.body.clientWidth");//documentElement.scrollWidth");
                    long viewportHeight1 = (long)((IJavaScriptExecutor)DriverDoSelenium).ExecuteScript("return window.innerHeight");//documentElement.scrollWidth");

                    int viewportWidth = (int)viewportWidth1;
                    int viewportHeight = (int)viewportHeight1;


                    // Split the Screen in multiple Rectangles
                    List<Rectangle> rectangles = new List<Rectangle>();
                    // Loop until the Total Height is reached
                    for (int i = 0; i < totalHeight; i += viewportHeight)
                    {
                        int newHeight = viewportHeight;
                        // Fix if the Height of the Element is too big
                        if (i + viewportHeight > totalHeight)
                        {
                            newHeight = totalHeight - i;
                        }
                        // Loop until the Total Width is reached
                        for (int ii = 0; ii < totalWidth; ii += viewportWidth)
                        {
                            int newWidth = viewportWidth;
                            // Fix if the Width of the Element is too big
                            if (ii + viewportWidth > totalWidth)
                            {
                                newWidth = totalWidth - ii;
                            }

                            // Create and add the Rectangle
                            Rectangle currRect = new Rectangle(ii, i, newWidth, newHeight);
                            rectangles.Add(currRect);
                        }
                    }

                    // Build the Image
                    stitchedImage = new Bitmap(totalWidth, totalHeight);
                    // Get all Screenshots and stitch them together
                    Rectangle previous = Rectangle.Empty;
                    foreach (var rectangle in rectangles)
                    {
                        // Calculate the Scrolling (if needed)
                        if (previous != Rectangle.Empty)
                        {
                            int xDiff = rectangle.Right - previous.Right;
                            int yDiff = rectangle.Bottom - previous.Bottom;
                            // Scroll
                            //selenium.RunScript(String.Format("window.scrollBy({0}, {1})", xDiff, yDiff));
                            ((IJavaScriptExecutor)DriverDoSelenium).ExecuteScript(String.Format("window.scrollBy({0}, {1})", xDiff, yDiff));
                            System.Threading.Thread.Sleep(200);
                        }

                        // Take Screenshot
                        var screenshot = ((ITakesScreenshot)DriverDoSelenium).GetScreenshot();

                        // Build an Image out of the Screenshot
                        Image screenshotImage;
                        using (MemoryStream memStream = new MemoryStream(screenshot.AsByteArray))
                        {
                            screenshotImage = Image.FromStream(memStream);
                        }

                        // Calculate the Source Rectangle
                        Rectangle sourceRectangle = new Rectangle(viewportWidth - rectangle.Width, viewportHeight - rectangle.Height, rectangle.Width, rectangle.Height);

                        // Copy the Image
                        using (Graphics g = Graphics.FromImage(stitchedImage))
                        {
                            g.DrawImage(screenshotImage, rectangle, sourceRectangle, GraphicsUnit.Pixel);
                        }

                        // Set the Previous Rectangle
                        previous = rectangle;
                    }

                    stitchedImage.Save(filename, ImageFormat.Png);
                }
                else
                {
                    Screenshot ss = ((ITakesScreenshot)DriverDoSelenium).GetScreenshot();
                    ss.SaveAsFile(filename, ScreenshotImageFormat.Png);
                }
                return filename;
            }
            catch (Exception ex)
            {
                Print("Erro ao printar a imagem.", ex);
                return null;
                // handle
            }

        }



        //////////////////////////////////////////////////////ARQUIVO DE CRIPTOGRAFIA////////////////////////////////////////////////////
        private static string ArrayBytesToHexString(byte[] conteudo) =>
            string.Concat(Array.ConvertAll<byte, string>(conteudo, b => b.ToString("X2")));

        private static Rijndael CriarInstanciaRijndael(string chave, string vetorInicializacao)
        {
            if ((chave == null) || (((chave.Length != 0x10) && (chave.Length != 0x18)) && (chave.Length != 0x20)))
            {
                return null;
            }
            if ((vetorInicializacao == null) || (vetorInicializacao.Length != 0x10))
            {
                return null;
            }
            Rijndael rijndael = Rijndael.Create();
            rijndael.Key = Encoding.ASCII.GetBytes(chave);
            rijndael.IV = Encoding.ASCII.GetBytes(vetorInicializacao);
            return rijndael;
        }

        public static string DecodeFrom64(string encodedData)
        {
            byte[] bytes = Convert.FromBase64String(encodedData);
            return Encoding.ASCII.GetString(bytes);
        }

        public static string Decriptar(string textoEncriptado)
        {
            string str2;
            if (string.IsNullOrWhiteSpace(textoEncriptado))
            {
                return "O conte\x00fado a ser decriptado n\x00e3o pode ser uma string vazia.";
            }
            if ((textoEncriptado.Length % 2) != 0)
            {
                return "O conte\x00fado a ser decriptado \x00e9 inv\x00e1lido.";
            }
            using (Rijndael rijndael = CriarInstanciaRijndael("ASDFGTRE$#@!6+|@", "QWERTYUIOP!@#+4+"))
            {
                ICryptoTransform transform = rijndael.CreateDecryptor(rijndael.Key, rijndael.IV);
                string str = null;
                try
                {
                    using (MemoryStream stream = new MemoryStream(HexStringToArrayBytes(textoEncriptado)))
                    {
                        using (CryptoStream stream2 = new CryptoStream(stream, transform, CryptoStreamMode.Read))
                        {
                            using (StreamReader reader = new StreamReader(stream2))
                            {
                                try
                                {
                                    str = reader.ReadToEnd();
                                }
                                catch (Exception)
                                {
                                }
                            }
                        }
                    }
                    str2 = str;
                }
                catch (Exception)
                {
                    str2 = "";
                }
            }
            return str2;
        }

        public static string EncodeTo64(string toEncode) =>
            Convert.ToBase64String(Encoding.ASCII.GetBytes(toEncode));

        public static string Encriptar(string textoNormal)
        {
            string str;
            if (string.IsNullOrWhiteSpace(textoNormal))
            {
                return "O conte\x00fado a ser encriptado n\x00e3o pode ser uma string vazia.";
            }
            using (Rijndael rijndael = CriarInstanciaRijndael("ASDFGTRE$#@!6+|@", "QWERTYUIOP!@#+4+"))
            {
                ICryptoTransform transform = rijndael.CreateEncryptor(rijndael.Key, rijndael.IV);
                using (MemoryStream stream = new MemoryStream())
                {
                    using (CryptoStream stream2 = new CryptoStream(stream, transform, CryptoStreamMode.Write))
                    {
                        using (StreamWriter writer = new StreamWriter(stream2))
                        {
                            writer.Write(textoNormal);
                        }
                    }
                    str = ArrayBytesToHexString(stream.ToArray());
                }
            }
            return str;
        }

        private static byte[] HexStringToArrayBytes(string conteudo)
        {
            try
            {
                int num = conteudo.Length / 2;
                byte[] buffer = new byte[num];
                for (int i = 0; i < num; i++)
                {
                    buffer[i] = Convert.ToByte(conteudo.Substring(i * 2, 2), 0x10);
                }
                return buffer;
            }
            catch (Exception)
            {
                return null;
            }
        }
        //////////////////////////////////////////////////////FIM ARQUIVO DE CRIPTOGRAFIA///////////////////////////////////////////////


    }


}



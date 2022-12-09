using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Text;
using System.Collections;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Drawing;
using ClosedXML.Excel;

namespace project
{
    public partial class WebForm1 : System.Web.UI.Page
    {



        public class PubConstant
        {
            /// <summary>
            /// 獲取連線字串
            /// </summary>
            public static string ConnectionString
            {
                get
                {


                    //string _connectionString = @"Server=LAPTOP-D4PBFKD3\SQLEXPRESS;Initial Catalog=ExcelData; Integrated Security=SSPI;";
                    string _connectionString = @"Server = localhost\SQLEXPRESS; Database = ExcelData; Trusted_Connection = True;";
                    return _connectionString;
                }
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            Label4.Text = "";           
            // if (!IsPostBack)
                Check_db();
                 
        }
        
        protected void Check_db()
        {

            using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(PubConstant.ConnectionString))
            {
                try
                {
                    string strSql = @"SELECT COUNT(name) FROM sys.Tables";
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;                              
                    object TCount = command.ExecuteScalar();
                    if (Convert.ToInt32(TCount) > 0)                
                    {
                        Label5.Text = "▲ 目前有資料儲存 ▲";
                        Label5.ForeColor = Color.Red;
                    }
                    else
                    {
                        Label5.Text = "○ 可正常使用 ○";
                        Label5.ForeColor = Color.Black;
                    }
                    sqlconn.Close();
                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        public class Global
        {
            public static Boolean check = false;
            public static ArrayList column = new ArrayList();
            public static ArrayList column2 = new ArrayList();
            public static int num;
            public static int num_g1;
            public static string same_col;
            public static ArrayList dif_col = new ArrayList();
            
        }

        protected void Button1_Click(object sender, EventArgs e)
        {

            ExcelUpload();
            Check_db();

        }
        protected void Button2_Click(object sender, EventArgs e)
        {
            if (Global.check)
            {
                ExcelUpload2();
            }
            else
            {
                MessageBox.Show("請先匯入主檔案，再匯入對比檔案!謝謝!");
            }
        }
        protected void ExcelUpload()
        {
            //存放檔案路徑
            String filepath = "";

            //存放副檔名
            string fileExtName = "";
            //檔名
            string mFileName = "";
            //伺服器上的相對路徑
            string mPath = "";

            if (FileUpload1.PostedFile.FileName != "")
            {
                //取得檔案路徑
                filepath = FileUpload1.PostedFile.FileName;
                //取得副檔名
                fileExtName = filepath.Substring(filepath.LastIndexOf(".") + 1);
                //取得伺服器上的相對路徑
                mPath = this.Request.PhysicalApplicationPath + "UpLoadFiles\\Excel\\";
                //取得檔名
                mFileName = filepath.Substring(filepath.LastIndexOf("\\") + 1);
                //儲存檔案到指定目錄
                if (!Directory.Exists(mPath))
                {
                    try
                    {
                        Directory.CreateDirectory(mPath);
                    }
                    catch
                    {
                        MessageBox.Show("伺服器建立存放目錄失敗");
                    }
                }

                //如果檔案已經存在則刪除原來的檔案
                if (File.Exists(mPath + mFileName))
                {
                    try
                    {
                        File.Delete(mPath + mFileName);
                    }
                    catch(Exception e)
                    {
                        MessageBox.Show("伺服器上存在相同檔案，刪除失敗。"+e.Message);
                    }
                }

                #region 判斷副檔名

                //判斷上傳檔案格式
                Boolean fileOK = false;

                if (FileUpload1.HasFile)
                {

                    String fileExtension = System.IO.Path.GetExtension(FileUpload1.FileName).ToLower();

                    String[] allowedExtensions = { ".xlsx" };

                    for (int i = 0; i < allowedExtensions.Length; i++)
                    {

                        if (fileExtension == allowedExtensions[i])
                        {

                            fileOK = true;

                        }

                    }

                }

                #endregion

                #region 判斷檔案是否上傳成功

                //判斷檔案是否上傳成功
                bool fileUpOK = false;
                if (fileOK)
                {
                    try
                    {
                        //檔案上傳到伺服器
                        FileUpload1.PostedFile.SaveAs(mPath + mFileName);

                        fileUpOK = true;
                    }
                    catch(Exception e)
                    {
                        MessageBox.Show("檔案上傳失敗！請確認檔案內容格式符合要求！"+e.Message);
                    }
                }
                else
                {

                    MessageBox.Show("上傳檔案的格式錯誤，應為.xlsx 格式！");

                }
                #endregion
                if (fileUpOK)
                {
                    System.Data.DataTable dt_User = new System.Data.DataTable();
                    try
                    {
                        //獲取Excel表中的內容
                        dt_User = GetList(mPath + mFileName);
                        GridView1.DataSource = dt_User;
                        GridView1.DataBind();
                        //this.GridView1.Style.Add("display", "none");
                        if (dt_User == null)
                        {
                            MessageBox.Show("獲取Excel內容失敗！");
                            return;
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("獲取Excel內容失敗！" + e.Message);
                    }
                    int rowNum = 0;
                    try
                    {
                        rowNum = dt_User.Rows.Count;
                    }
                    catch
                    {
                        MessageBox.Show("Excel表獲取失敗！");
                    }
                    if (rowNum == 0)
                    {
                        MessageBox.Show("Excel為空表，無資料！");
                    }
                    else
                    {
                        Global.check = true;
                        //資料儲存
                        SaveToDataBase(dt_User, 0);
                        LoadTable(GridView1, 0);
                        Label1.Text = mFileName+" 匯入成功";
                        Label1.ForeColor = Color.Red;
                        Label1.Font.Bold = true;
                    }
                }
            }
            else
            {
                MessageBox.Show("請先選擇匯入檔案後，再執行匯入！謝謝！");
            }
        }

        protected void ExcelUpload2()
        {
            //存放檔案路徑
            String filepath = "";

            //存放副檔名
            string fileExtName = "";
            //檔名
            string mFileName = "";
            //伺服器上的相對路徑
            string mPath = "";
            
            if (FileUpload2.PostedFile.FileName != "")
            {
                //取得檔案路徑
                filepath = FileUpload2.PostedFile.FileName;
                //取得副檔名
                fileExtName = filepath.Substring(filepath.LastIndexOf(".") + 1);
                //取得伺服器上的相對路徑
                mPath = this.Request.PhysicalApplicationPath + "UpLoadFiles\\Excel_cmp\\";
                //取得檔名
                mFileName = filepath.Substring(filepath.LastIndexOf("\\") + 1);
                //儲存檔案到指定目錄
                if (!Directory.Exists(mPath))
                {
                    try
                    {
                        Directory.CreateDirectory(mPath);
                    }
                    catch
                    {
                        MessageBox.Show("伺服器建立存放目錄失敗");
                    }
                }

                //如果檔案已經存在則刪除原來的檔案
                if (File.Exists(mPath + mFileName))
                {
                    try
                    {
                        File.Delete(mPath + mFileName);
                    }
                    catch
                    {
                        MessageBox.Show("伺服器上存在相同檔案，刪除失敗。");
                    }
                }

                #region 判斷副檔名

                //判斷上傳檔案格式
                Boolean fileOK = false;

                if (FileUpload2.HasFile)
                {

                    String fileExtension = System.IO.Path.GetExtension(FileUpload2.FileName).ToLower();

                    String[] allowedExtensions = { ".xlsx" };

                    for (int i = 0; i < allowedExtensions.Length; i++)
                    {

                        if (fileExtension == allowedExtensions[i])
                        {

                            fileOK = true;

                        }

                    }

                }

                #endregion

                #region 判斷檔案是否上傳成功

                //判斷檔案是否上傳成功
                bool fileUpOK = false;
                if (fileOK)
                {
                    try
                    {
                        //檔案上傳到伺服器
                        FileUpload2.PostedFile.SaveAs(mPath + mFileName);

                        fileUpOK = true;
                    }
                    catch
                    {
                        MessageBox.Show("檔案上傳失敗！請確認檔案內容格式符合要求！");
                    }
                }
                else
                {

                    MessageBox.Show("上傳檔案的格式錯誤，應為.xlsx 格式！");

                }
                #endregion
                if (fileUpOK)
                {
                    System.Data.DataTable dt_User = new System.Data.DataTable();
                    try
                    {
                        //獲取Excel表中的內容
                        dt_User = GetList(mPath + mFileName);
                        GridView2.DataSource = dt_User;
                        GridView2.DataBind();
                        this.GridView2.Style.Add("display", "none");
                        if (dt_User == null)
                        {
                            MessageBox.Show("獲取Excel內容失敗！");
                            return;
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("獲取Excel內容失敗！" + e.Message);
                    }
                    int rowNum = 0;
                    try
                    {
                        rowNum = dt_User.Rows.Count;
                    }
                    catch
                    {
                        MessageBox.Show("Excel表獲取失敗！");
                    }
                    if (rowNum == 0)
                    {
                        MessageBox.Show("Excel為空表，無資料！");
                    }
                    else
                    {
                        //資料儲存
                        SaveToDataBase(dt_User, 1);
                        LoadTable(GridView2, 1);
                        Label2.Text = mFileName+"匯入成功";
                        Label2.ForeColor = Color.Red;
                        Label2.Font.Bold = true;

                    }
                }
            }
            else
            {
                MessageBox.Show("請先選擇匯入檔案後，再執行匯入！謝謝！");
            }
        }

        public System.Data.DataTable GetList(string FilePath)
        {
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'";
            string strSql = string.Empty;

            //string workSheetName = Get_FistWorkBookName(FilePath);
            //第一個工作表的名稱。考慮到穩定性，就直接寫死了。
            string sheetName = "工作表1";
            if (sheetName != "")
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    // Sheet名稱，必需用中括號 [] 包起來
                    string TSQL = "SELECT * FROM [" + sheetName + "$]";
                    OleDbDataAdapter da = new OleDbDataAdapter(TSQL, connectionString);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    return dt;
                }
            }
            else
            {
                return null;
            }
        }

        protected void SaveToDataBase(System.Data.DataTable dt_user, int n)
        {
            string sheetName;
            if (n == 0) { sheetName = "工作表1"; }
            else { sheetName = "工作表2"; }
            string strSql = string.Format("if not exists(select * from sysobjects where name = '{0}') create table {0}(", sheetName);   //以sheetName為表名   
            foreach (System.Data.DataColumn c in dt_user.Columns)
            {
                strSql += string.Format("[{0}] varchar(255),", c.ColumnName);
            }
            strSql = strSql.Trim(',') + ")";
            strSql+= string.Format(" else truncate table {0}", sheetName);
            /*if (n == 0)
            {
                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(PubConstant.ConnectionString))
                {
                    try
                    {
                        strSql += @" insert into dbo.工作表1 ( ";
                        
                        for (int j = 0; j < GridView1.Rows.Count; j++)
                        {
                            for (int i = 0; i < Global.num_g1; i++)
                            {

                                if (i != Global.num_g1 - 1)
                                    strSql += Global.column[i].ToString().Trim() + ", ";
                                else
                                    strSql += Global.column[i].ToString().Trim() + ") values (N'";                               
                            }
                            for(int k=0;k < Global.num_g1; k++)
                            {
                               strSql += GridView1.Rows[j].Cells[k].ToString().Trim()+ "', ";
                            }
                            
                        }
                        
                        sqlconn.Open();
                        System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                        command.CommandText = strSql;
                        command.ExecuteNonQuery();
                        sqlconn.Close();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            else
            {


            }*/
            try
            {
                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(PubConstant.ConnectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                }
                //用bcp匯入資料      
                //excel檔案中列的順序必須和資料表的列順序一致，因為資料匯入時，是從excel檔案的第二行資料開始，不管資料表的結構是什麼樣的，反正就是第一列的資料會插入到資料表的第一列欄位中，第二列的資料插入到資料表的第二列欄位中，以此類推，它本身不會去判斷要插入的資料是對應資料表中哪一個欄位的   
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(PubConstant.ConnectionString))
                {
                    //bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(bcp_SqlRowsCopied);
                    bcp.BatchSize = 100;//每次傳輸的行數      
                    bcp.NotifyAfter = 100;//進度提示的行數      
                    bcp.DestinationTableName = sheetName;//目標表      
                    bcp.WriteToServer(dt_user);
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




        protected DataTable Query_Same()
        {

            string strSql = @"select a.* from  dbo.工作表1 a 
                              inner join dbo.工作表2 b on";

            //strSql += "a." + Global.column2[0].ToString() + "= b." + Global.column2[0].ToString();
            for (int i = 0; i < Global.num; i++)
            {
                if (i != Global.num - 1)
                    strSql += " a.[" + Global.column2[i].ToString().Trim() + "] = b.[" + Global.column2[i].ToString().Trim() + "] and ";
                else
                    strSql += " a.[" + Global.column2[i].ToString().Trim() + "] = b.[" + Global.column2[i].ToString().Trim()+"]";

            }

            DataTable dt = new DataTable();
            try
            {
                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(PubConstant.ConnectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    SqlDataAdapter da = new SqlDataAdapter(command);
                    da.Fill(dt);

                    da.Dispose();
                    sqlconn.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return dt;
        }
        protected DataTable Query_Dif()
        {
            string strSql = @"select ";
            for(int i = 0; i < Global.num; i++)
            {
                if( i!=Global.num - 1)
                    strSql += @" a.["+ Global.column2[i].ToString().Trim()+"] ," ;
                else
                    strSql += @" a.[" + Global.column2[i].ToString().Trim() + "]";
            }
            
            strSql += " into [##Table1] from dbo.工作表1 a except";
            strSql += " select ";
            for (int j = 0; j < Global.num; j++)
            {
                if (j != Global.num - 1)
                    strSql += @" b.[" + Global.column2[j].ToString().Trim() + "] ,";
                else
                    strSql += @" b.[" + Global.column2[j].ToString().Trim() + "]";
            }
            strSql += " from dbo.工作表2 b";
            strSql += " select  a.* from dbo.工作表1 a inner join ##Table1 b on ";
            for (int i = 0; i < Global.num; i++)
            {
                if (i != Global.num - 1)
                    strSql += " a.[" + Global.column2[i].ToString().Trim() + "] = b.[" + Global.column2[i].ToString().Trim() + "] and ";
                else
                    strSql += " a.[" + Global.column2[i].ToString().Trim() + "] = b.[" + Global.column2[i].ToString().Trim() + "]";

            }
            strSql += " drop table ##Table1 ";
            /*
            string strSql = @"select a.* from  dbo.工作表1 a 
                              where ";
            for(int i =0;i< Global.num; i++)
            {
                if (i != Global.num - 1)
                    strSql += " a.[" + Global.column2[i].ToString().Trim() + "] not in (select b.[" + Global.column2[i].ToString().Trim() + "] from dbo.工作表2 b) or";
                else
                    strSql += " a.[" + Global.column2[i].ToString().Trim() + "] not in (select b.[" + Global.column2[i].ToString().Trim() + "] from dbo.工作表2 b) ";
            }          
            for (int i = 0; i < Global.num; i++)
            {
                if (i != Global.num - 1)
                    strSql += " a." + Global.column2[i].ToString().Trim() + " = b." + Global.column2[i].ToString().Trim() + " and ";
                else
                    strSql += " a." + Global.column2[i].ToString().Trim() + " = b." + Global.column2[i].ToString().Trim();

            }
            strSql += " where ";
            for(int j = 0; j < Global.num; j++)
            {
                if (j != Global.num - 1)
                    strSql += "a." + Global.column2[j].ToString().Trim() + " is null " + "or b."+ Global.column2[j].ToString().Trim()+" is null and "; 
                else
                    strSql += "a." + Global.column2[j].ToString().Trim() + " is null " + "or b." + Global.column2[j].ToString().Trim() + " is null";
            }*/
            DataTable dt = new DataTable();
            try
            {
                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(PubConstant.ConnectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    SqlDataAdapter da = new SqlDataAdapter(command);
                    da.Fill(dt);

                    da.Dispose();
                    sqlconn.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }          
            return dt;
        }
        protected DataTable Query_Combine()
        {
            //check same column
            string[] arr = (string[])Global.column.ToArray(Type.GetType("System.String"));
            string[] arr2 = (string[])Global.column2.ToArray(Type.GetType("System.String"));
            for (int i = 0; i < Global.num_g1; i++)
            {
                for(int j = 0; j < Global.num; j++)
                {
                    if(arr[i] == arr2[j])
                    {
                        Global.same_col = arr[i];
                    }
                }
            }
            for(int i = 0; i < Global.num; i++)
            {
                if( arr2[i] != Global.same_col)
                {
                    Global.dif_col.Add(arr2[i]);
                }
            }
            
            string strSql = "select a.*, ";
            foreach (string i in Global.dif_col)
            {
                strSql += "b.[" + i + "] ,";
            }
            strSql = strSql.TrimEnd(',');
            strSql += @" from  dbo.工作表1 a 
                         inner join dbo.工作表2 b on a.["+ Global.same_col + "] = "+ "b.["+ Global.same_col+"]";

            DataTable dt = new DataTable();
            try
            {
                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(PubConstant.ConnectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    SqlDataAdapter da = new SqlDataAdapter(command);
                    da.Fill(dt);

                    da.Dispose();
                    sqlconn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return dt;

        }
        protected void LoadTable(GridView gridView, int n)
        {
            int num = gridView.HeaderRow.Cells.Count;
            if (n == 0)
            {
                for (int i = 0; i < num; i++)
                {
                    Global.column.Add(gridView.HeaderRow.Cells[i].Text);
                }
                Global.num_g1 = gridView.HeaderRow.Cells.Count;
            }
            else
            {
                for (int i = 0; i < num; i++)
                {
                    Global.column2.Add(gridView.HeaderRow.Cells[i].Text);
                }
                Global.num = gridView.HeaderRow.Cells.Count;
                

            }
        }

        public class MessageBox
        {
            private static Hashtable m_executingPages = new Hashtable();

            private MessageBox() { }
            /// <summary>
            /// MessageBox訊息窗
            /// </summary>
            /// <param name="sMessage">要顯示的訊息</param>
            public static void Show(string sMessage)
            {
                // If this is the first time a page has called this method then
                if (!m_executingPages.Contains(HttpContext.Current.Handler))
                {
                    // Attempt to cast HttpHandler as a Page.
                    Page executingPage = HttpContext.Current.Handler as Page;

                    if (executingPage != null)
                    {
                        // Create a Queue to hold one or more messages.
                        Queue messageQueue = new Queue();

                        // Add our message to the Queue
                        messageQueue.Enqueue(sMessage);

                        // Add our message queue to the hash table. Use our page reference
                        // (IHttpHandler) as the key.
                        m_executingPages.Add(HttpContext.Current.Handler, messageQueue);

                        // Wire up Unload event so that we can inject some JavaScript for the alerts.
                        executingPage.Unload += new EventHandler(ExecutingPage_Unload);
                    }
                }
                else
                {
                    // If were here then the method has allready been called from the executing Page.
                    // We have allready created a message queue and stored a reference to it in our hastable. 
                    Queue queue = (Queue)m_executingPages[HttpContext.Current.Handler];

                    // Add our message to the Queue
                    queue.Enqueue(sMessage);
                }
            }


            // Our page has finished rendering so lets output the JavaScript to produce the alert's
            private static void ExecutingPage_Unload(object sender, EventArgs e)
            {
                // Get our message queue from the hashtable
                Queue queue = (Queue)m_executingPages[HttpContext.Current.Handler];

                if (queue != null)
                {
                    StringBuilder sb = new StringBuilder();

                    // How many messages have been registered?
                    int iMsgCount = queue.Count;

                    // Use StringBuilder to build up our client slide JavaScript.
                    sb.Append("<script language='javascript'>");

                    // Loop round registered messages
                    string sMsg;
                    while (iMsgCount-- > 0)
                    {
                        sMsg = (string)queue.Dequeue();
                        //sMsg = sMsg.Replace( "\n", "\\n" ); //這部分是我mark掉的
                        sMsg = sMsg.Replace("\"", "'");

                        //W3c建議要避開的危險字元
                        //&;`'\"|*?~<>^()[]{}$\n\r
                        sMsg = sMsg.Replace("\n", "_");
                        sMsg = sMsg.Replace("\r", "_");

                        sb.Append(@"alert( """ + sMsg + @""" );");
                    }

                    // Close our JS
                    sb.Append(@"</script>");

                    // Were done, so remove our page reference from the hashtable
                    m_executingPages.Remove(HttpContext.Current.Handler);

                    // Write the JavaScript to the end of the response stream.
                    HttpContext.Current.Response.Write(sb.ToString());
                }
            }
        }

        protected void Button3_Click(object sender, EventArgs e)
        {
            
            
            DataTable dt = new DataTable();
            if (RadioButton1.Checked) dt = Query_Same();
            if (RadioButton2.Checked) dt = Query_Dif();
            if (RadioButton3.Checked) dt = Query_Combine();
            SaveToExcel(dt);
            //GridView3.DataSource = dt;
            //GridView3.DataBind();
            /*XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "工作表1");


            // Prepare the response
            //Response.Charset = "BIG5";
            //Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5");           
            HttpResponse httpResponse = Response;         
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"對比結果.xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                wb.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);               
                memoryStream.Close();              
            }
            
            httpResponse.End();
            */
        }
        protected void SaveToExcel(DataTable dt)
        {
            
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "工作表1");
            // Prepare the response
            //Response.Charset = "BIG5";
            //Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5");           
            HttpResponse httpResponse = Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"對比結果.xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                wb.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();
        }

        protected void Button4_Click(object sender, EventArgs e)
        {
            using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(PubConstant.ConnectionString))
            {
                try
                {
                    string strSql = @"drop table if exists dbo.工作表1,dbo.工作表2";
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                    Label1.Text = "";
                    Label2.Text = "";
                    Label4.Text = "清除成功";
                    Global.column.Clear();
                    Global.column2.Clear();
                    Global.same_col = "";
                    Global.dif_col.Clear();

                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Check_db();
            }
        }
        protected void Button5_Click(object sender, EventArgs e)
        {
            using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(PubConstant.ConnectionString))
            {
                try
                {
                    string strSql = @"DROP TABLE if exists dbo.工作表2";
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();                   
                    Label2.Text = "";
                    Label4.Text = "清除成功";
                    Global.column2.Clear();                   
                    Global.same_col = "";
                    Global.dif_col.Clear();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            
             
        }


    }
}
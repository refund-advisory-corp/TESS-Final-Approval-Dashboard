using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using AxAcroPDFLib;
using AcroBrokerLib;
using Acrobat;
using AcroPDFLib;
using AcrobatAccessLib;
using System.DirectoryServices.AccountManagement;
using ActiveDs;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System.IO;

namespace TESS_Dashboard
{
    public partial class Main_Window : Form
    {
        public string ExMessage;
        public int OnRecord = 0;

        public Main_Window()
        {
            InitializeComponent();
        }

        private void Main_Window_Load(object sender, EventArgs e)
        {
            SqlConnection CONN = new SqlConnection(ConfigurationManager.ConnectionStrings["TESS Dashboard.Properties.Settings.REFTBL4ConnectionString"].ConnectionString);
            CONN.Open();
            try
            {
                GetCountyOptions(CONN);
                GetClientStatusOptions(CONN);
                GetSalesStatusOptions(CONN);

                //MPA 12/18/2019 Here, lock final approval if user is not privileged!
                PrincipalContext pc = new PrincipalContext(ContextType.Domain, "corp.kwardcpa.com");
                string username = "";
                username = Environment.UserName;

                UserPrincipal user = UserPrincipal.FindByIdentity(pc, username);

                GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, "Executives");
                
                CheckBox_ASLS_STATUS.Enabled = !string.IsNullOrEmpty(username) && user.IsMemberOf(group);

            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error in Main_Window_Load: " + Ex.Message + " ... " + Ex.StackTrace);
            }
            finally
            {
                CONN.Close();
            }
            
        }

        private void GetClientStatusOptions(SqlConnection con)
        {
            string SqlStr = "SELECT DISTINCT ACLT_STATUS FROM tblACCT_CLIENT WHERE ACLT_STATUS LIKE 'Need%' ORDER BY ACLT_STATUS";

            SqlCommand cmd = new SqlCommand(SqlStr, con);

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);

            ComboBox_NeedsReview.DisplayMember = "ACLT_STATUS";

            ComboBox_NeedsReview.DataSource = dt;
        }

        private void GetCountyOptions(SqlConnection con)
        {
            string SqlStr = "SELECT CNTY_ABBR FROM tblCOUNTY UNION Select 'All' ORDER BY CNTY_ABBR";
            SqlCommand cmd = new SqlCommand(SqlStr, con);

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            ComboBox_County.DisplayMember = "CNTY_ABBR";

            ComboBox_County.DataSource = dt;
        }

        private void GetSalesStatusOptions(SqlConnection con)
        {
            string SqlStr = "SELECT DISTINCT ASLS_STATUS FROM tblACCT_SALES UNION Select 'All' ORDER BY ASLS_STATUS";
            SqlCommand cmd = new SqlCommand(SqlStr, con);

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            ComboBox_ASLS_STATUS.DisplayMember = "ASLS_STATUS";
            
            ComboBox_ASLS_STATUS.DataSource = dt;
        }


        private bool SearchRecords(SqlConnection con, int acct = 0)
        {

            Dictionary<string, System.Type> RequestFields = new Dictionary<string, System.Type>(); //MPA 12/13/2019 organize field getting
            RequestFields.Add("ACCT", typeof(int));

            RequestFields.Add("ACCT_ALIAS", typeof(string));
            RequestFields.Add("ACCT_TAX_YEAR", typeof(int));
            RequestFields.Add("ACCT_REFUND", typeof(decimal));
            RequestFields.Add("ACCT_CANCEL", typeof(bool));

            RequestFields.Add("ACCT_CONTACT_NAME1_TITLE", typeof(string));
            RequestFields.Add("ACCT_CONTACT_NAME1_FIRST", typeof(string));
            RequestFields.Add("ACCT_CONTACT_NAME1_MIDDLE", typeof(string));
            RequestFields.Add("ACCT_CONTACT_NAME1_LAST", typeof(string));
            RequestFields.Add("ACCT_CONTACT_NAME1_SUFFIX", typeof(string));

            RequestFields.Add("ACCT_CONTACT_NAME2_TITLE", typeof(string));
            RequestFields.Add("ACCT_CONTACT_NAME2_FIRST", typeof(string));
            RequestFields.Add("ACCT_CONTACT_NAME2_MIDDLE", typeof(string));
            RequestFields.Add("ACCT_CONTACT_NAME2_LAST", typeof(string));
            RequestFields.Add("ACCT_CONTACT_NAME2_SUFFIX", typeof(string));

            RequestFields.Add("ASLS_STATUS", typeof(string));
            RequestFields.Add("ASLS_SALESMAN", typeof(string));

            //RequestFields.Add("ACCT_CONTACT_ADDR", "tblACCT");
            //RequestFields.Add("ACCT_CONTACT_CITY", "tblACCT");
            //RequestFields.Add("ACCT_CONTACT_STATE", "tblACCT");
            //RequestFields.Add("ACCT_CONTACT_ZIP", "tblACCT");
            RequestFields.Add("ACCT_EX_HEX_GOAL", typeof(bool));
            RequestFields.Add("ACCT_EX_O65_GOAL", typeof(bool));
            RequestFields.Add("ACCT_EX_DP_GOAL", typeof(bool));
            RequestFields.Add("ACCT_EX_GOAL_PCT", typeof(decimal));
                
            RequestFields.Add("PRCL_PID", typeof(string));
            RequestFields.Add("PRCL_SITUS_ADDR", typeof(string));
            RequestFields.Add("PRCL_SITUS_CITY", typeof(string));
            RequestFields.Add("PRCL_SITUS_ZIP", typeof(string));
            RequestFields.Add("PRCL_SITUS_STATE", typeof(string));

            RequestFields.Add("ACLT_APP_APPROVED", typeof(DateTime));
            RequestFields.Add("ACLT_SIGNED", typeof(bool));
            RequestFields.Add("ACLT_DOB1", typeof(DateTime));
            RequestFields.Add("ACLT_DOB2", typeof(DateTime));
            RequestFields.Add("ACLT_FEE_PCT", typeof(decimal));
            RequestFields.Add("ACLT_EX_DP_APPLIED", typeof(bool));
            RequestFields.Add("ACLT_EX_HEX_APPLIED", typeof(bool));
            RequestFields.Add("ACLT_EX_O65_APPLIED", typeof(bool));
            RequestFields.Add("ACLT_FILE_PREV", typeof(bool));
            RequestFields.Add("ACLT_FILE_LAST", typeof(bool));
            RequestFields.Add("ACLT_FILE_CURR", typeof(bool));
            RequestFields.Add("ACLT_FILE_NEXT", typeof(bool));
            RequestFields.Add("ACLT_OWNER1_SIGNED", typeof(bool));
            RequestFields.Add("ACLT_OWNER2_SIGNED", typeof(bool));
            RequestFields.Add("ACLT_STATUS", typeof(string));
            RequestFields.Add("ACLT_MISC_NOTES", typeof(string));
                
            RequestFields.Add("POWN_DEED_DATE", typeof(DateTime)); 
            RequestFields.Add("POWN_OWNER_NAME1", typeof(string));
            RequestFields.Add("POWN_OWNER_NAME2", typeof(string));
            RequestFields.Add("POWN_MAIL_ADDR1", typeof(string));
            RequestFields.Add("POWN_MAIL_ADDR2", typeof(string));
            RequestFields.Add("POWN_MAIL_CITY", typeof(string));
            RequestFields.Add("POWN_MAIL_STATE", typeof(string));
            RequestFields.Add("POWN_MAIL_ZIP", typeof(string));

            RequestFields.Add("(Case When ISNULL(ACCT_EX_HEX_GOAL,0) = 1 AND IsNull(ACCT_EX_O65_GOAL, 0) = 0 AND IsNull(ACCT_EX_DP_GOAL, 0) = 0 Then 1 Else 0 End) AS HexOnly", typeof(bool));
            RequestFields.Add("(Case When ISNULL([ACCT_EX_HEX_GOAL],0) = 1 AND IsNull([ACCT_EX_O65_GOAL], 0) = 1 Then 1 Else 0 End) AS HexAndO65", typeof(bool));
            
            SQLcommander RecordListCommand = new SQLcommander(con);
            RecordListCommand.TableName = "qryACCT_SALES_PARCEL_OWNER_CLIENT"; //"qryACCT_SALES_OWNER_CLIENT";

            foreach (string aKey in RequestFields.Keys)
            {
                RecordListCommand.AddRequest(aKey);
            }

            Dictionary<string, Tuple<string, object>> ConditionFields = new Dictionary<string, Tuple<string, object>>();
            ConditionFields.Add("ISNULL(ACLT_APP_APPROVED, 0) = 0", new Tuple<string, object>(null, null)); //Allow for false, too?

            if (!((string)(((System.Data.DataRowView)ComboBox_County.SelectedValue).Row.ItemArray[0]) == "All")) 
            {
                ConditionFields.Add("ACCT_CNTY_ABBR = @ACCT_CNTY_ABBR", new Tuple<string, object>("@ACCT_CNTY_ABBR", ((System.Data.DataRowView)ComboBox_County.SelectedValue).Row.ItemArray[0]));
            }

            if (CheckBox_ASLS_STATUS.Checked || (!CheckBox_ASLS_STATUS.Checked && !CheckBox_ASLS_STATUS_TESS.Checked && ((System.Data.DataRowView)ComboBox_ASLS_STATUS.SelectedValue).Row.ItemArray[0].ToString().Replace(" ", "") == "") )
            {
                ConditionFields.Add("ASLS_STATUS Like 'Lead Completed%'", new Tuple<string, object>(null, null));
            }
            else if (CheckBox_ASLS_STATUS_TESS.Checked)
            {
                ConditionFields.Add("ASLS_STATUS = 'Rep Claims Complete - TESS'", new Tuple<string, object>(null, null));
            }
            else if (!(((System.Data.DataRowView)ComboBox_ASLS_STATUS.SelectedValue).Row.ItemArray[0].ToString() == "All"))
            {
                ConditionFields.Add("ASLS_STATUS = '" + ((System.Data.DataRowView)ComboBox_ASLS_STATUS.SelectedValue).Row.ItemArray[0].ToString() + "'", new Tuple<string, object>(null, null));
            }
            

            if (DateTimePicker_App_Recv_From.Checked)
            {
                ConditionFields.Add("ACLT_APPL_RECV >= @ACLT_APPL_RECV_FROM", new Tuple<string, object>("@ACLT_APPL_RECV_FROM", new DateTime(DateTimePicker_App_Recv_From.Value.Year, DateTimePicker_App_Recv_From.Value.Month, DateTimePicker_App_Recv_From.Value.Day)));
            }

            if (DateTimePicker_App_Recv_To.Checked)
            {
                ConditionFields.Add("ACLT_APPL_RECV <= @ACLT_APPL_RECV_TO", new Tuple<string, object>("@ACLT_APPL_RECV_TO", new DateTime(DateTimePicker_App_Recv_To.Value.Year, DateTimePicker_App_Recv_To.Value.Month ,DateTimePicker_App_Recv_To.Value.Day)));
            }

            ConditionFields.Add("ACLT_AGR_TYPE = @Type", new Tuple<string, object>("@Type", 'T'));

            bool returnbool = false;

            if (acct == 0)
            {
                foreach (string aKey in ConditionFields.Keys)
                {
                    RecordListCommand.Conditions.Add(aKey);

                    if (ConditionFields[aKey].Item1 != null)
                    {
                        RecordListCommand.Parameters.Add(ConditionFields[aKey].Item1, ConditionFields[aKey].Item2);
                    }
                }

                RecordListCommand.Select();
                object[,] ReturnedRecords = RecordListCommand.SelectResults;

                DataGridView_Sales_Records.Columns.Clear();

                int counter = 0;
                foreach (string aKey in RequestFields.Keys)
                {
                    string ColumnName = aKey;
                    if (ColumnName.IndexOf(" AS ") > 0)
                    {
                        ColumnName = ColumnName.Substring(ColumnName.IndexOf(" AS ") + 4, ColumnName.Length - (ColumnName.IndexOf(" AS ") + 4));
                    }
                    DataGridView_Sales_Records.Columns.Add(ColumnName, ColumnName);

                    DataGridView_Sales_Records.Columns[counter++].ValueType = RequestFields[aKey];
                }

                for (int i = 0; i <= ReturnedRecords.GetUpperBound(0); i++)
                {

                    object[] RowToAdd = new object[ReturnedRecords.GetUpperBound(1) + 1];

                    for (int j = 0; j <= ReturnedRecords.GetUpperBound(1); j++)
                    {
                        dynamic ColumnType = DataGridView_Sales_Records.Columns[j].ValueType;
                        if (ColumnType == typeof(int))
                        {
                            RowToAdd[j] = ReturnedRecords[i, j];
                        }
                        else if (ColumnType == typeof(String) || ColumnType == typeof(string))
                        {
                            RowToAdd[j] = ReturnedRecords[i, j] ?? "";
                        }
                        else if (ColumnType == typeof(DateTime))
                        {
                            RowToAdd[j] = ReturnedRecords[i, j] ?? new DateTime(1, 1, 1);
                        }
                        else if (ColumnType == typeof(bool))
                        {
                            RowToAdd[j] = ReturnedRecords[i, j] ?? false;
                        }
                        else
                        {
                            RowToAdd[j] = ReturnedRecords[i, j];
                        }
                    }

                    DataGridView_Sales_Records.Rows.Add(RowToAdd);
                }

                ListSortDirection TheDirection = CheckBox_ReverseSort.Checked ? ListSortDirection.Ascending : ListSortDirection.Descending;

                DataGridView_Sales_Records.Sort(DataGridView_Sales_Records.Columns[DataGridView_Sales_Records.Columns["ACLT_FILE_PREV"].Index], TheDirection);
                DataGridView_Sales_Records.Sort(DataGridView_Sales_Records.Columns[DataGridView_Sales_Records.Columns["ACLT_FILE_LAST"].Index], TheDirection);
                DataGridView_Sales_Records.Sort(DataGridView_Sales_Records.Columns[DataGridView_Sales_Records.Columns["ACLT_FILE_CURR"].Index], TheDirection);
                DataGridView_Sales_Records.Sort(DataGridView_Sales_Records.Columns[DataGridView_Sales_Records.Columns["ACLT_FILE_NEXT"].Index], TheDirection);
                DataGridView_Sales_Records.Sort(DataGridView_Sales_Records.Columns[DataGridView_Sales_Records.Columns["HexOnly"].Index], TheDirection);
                DataGridView_Sales_Records.Sort(DataGridView_Sales_Records.Columns[DataGridView_Sales_Records.Columns["HexAndO65"].Index], TheDirection);

                returnbool = ReturnedRecords.GetUpperBound(0) > -1;
            }
            else
            {
                ConditionFields.Add("ACCT = @Acct", new Tuple<string, object>("@Acct", acct));

                foreach (string aKey in ConditionFields.Keys)
                {
                    RecordListCommand.Conditions.Add(aKey);

                    if (ConditionFields[aKey].Item1 != null)
                    {
                        RecordListCommand.Parameters.Add(ConditionFields[aKey].Item1, ConditionFields[aKey].Item2);
                    }
                }

                RecordListCommand.Select();
                object[,] ReturnedRecords = RecordListCommand.SelectResults;

                object[] RowToAdd = new object[ReturnedRecords.GetUpperBound(1) + 1];

                int searchValue = acct;
                int rowIndex = -1;
                foreach (DataGridViewRow row in DataGridView_Sales_Records.Rows)
                {
                    if (row.Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value.Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        break;
                    }
                }

                if (ReturnedRecords.GetUpperBound(0) > -1)
                {
                    for (int j = 0; j <= ReturnedRecords.GetUpperBound(1); j++)
                    {
                        dynamic ColumnType = DataGridView_Sales_Records.Columns[j].ValueType;
                        if (ColumnType == typeof(int))
                        {
                            RowToAdd[j] = ReturnedRecords[0, j];
                        }
                        else if (ColumnType == typeof(String) || ColumnType == typeof(string))
                        {
                            RowToAdd[j] = ReturnedRecords[0, j] ?? "";
                        }
                        else if (ColumnType == typeof(DateTime))
                        {
                            RowToAdd[j] = ReturnedRecords[0, j] ?? new DateTime(1, 1, 1);
                        }
                        else if (ColumnType == typeof(bool))
                        {
                            RowToAdd[j] = ReturnedRecords[0, j] ?? false;
                        }
                        else
                        {
                            RowToAdd[j] = ReturnedRecords[0, j];
                        }
                    }



                    DataGridView_Sales_Records.Rows[rowIndex].SetValues(RowToAdd);
                }

                returnbool = ReturnedRecords.GetUpperBound(0) > -1;
            }
            return returnbool;
        }


        private string AliasFromAcct(int Acct)
        {
            SQLcommander AliasCommand = new SQLcommander();
            AliasCommand.TableName = "tblACCT";
            AliasCommand.Requests.Add("ACCT_ALIAS");
            AliasCommand.Conditions.Add("ACCT = @acct");
            AliasCommand.Parameters.Add("@acct", Acct);
            AliasCommand.Select();
            return (string)AliasCommand.SelectResults[0,0];
        }

        private void ComboBox_Sales_Status_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void Search()
        {

            bool SearchRecordsResult;

            SqlConnection CONN = new SqlConnection(ConfigurationManager.ConnectionStrings["TESS Dashboard.Properties.Settings.REFTBL4ConnectionString"].ConnectionString);
            CONN.Open();
            try
            {
                SearchRecordsResult = SearchRecords(CONN);
            }
            catch (Exception Ex)
            {

                MessageBox.Show("Error in Search(): " + Ex.Message + " ... " + Ex.StackTrace);
                SearchRecordsResult = false;
            }
            finally
            {
                CONN.Close();
            }


            if (SearchRecordsResult)
            {
                OnRecord = 0;
                LoadRecord();
            }
            else
            {
                MessageBox.Show("No records found!");
            }
            
        }

        private void Button_Search_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void Button_Accept_Click(object sender, EventArgs e)
        {
            
            SqlConnection CONN = new SqlConnection(ConfigurationManager.ConnectionStrings["TESS Dashboard.Properties.Settings.REFTBL4ConnectionString"].ConnectionString);
            CONN.Open();
            try
            {
                if (CheckBox_ASLS_STATUS.Checked)
                {
                    SQLcommander ApproveCommand = new SQLcommander(CONN);
                    ApproveCommand.TableName = "tblACCT_CLIENT";
                    ApproveCommand.AddRequest("ACLT_APP_APPROVED = '" + DateTime.Now.ToString("yyyy-MM-dd'T'HH:mm:ss") + "'");
                    ApproveCommand.AddCondition("ACLT_ACCT = " + Convert.ToInt32(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value));
                    ApproveCommand.Update();

                    InsertCallRecord(CONN, Convert.ToInt32(Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value)), "Approved for submittal by " + GetInitials());
                }
                else if (CheckBox_ASLS_STATUS_TESS.Checked)
                {
                    SQLcommander ApproveCommand = new SQLcommander(CONN);
                    ApproveCommand.TableName = "tblACCT_SALES";
                    ApproveCommand.AddRequest("ASLS_STATUS = 'Lead Completed'");
                    ApproveCommand.AddCondition("ASLS_ACCT = " + Convert.ToInt32(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value));
                    ApproveCommand.Update();

                    InsertCallRecord(CONN, Convert.ToInt32(Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value)), "TESS forms passed initial review");
                }
                else
                {
                    MessageBox.Show("Unexpected contingency: Accept_Click! ASLS_STATUS not accounted for.");
                }

                if (OnRecord < DataGridView_Sales_Records.Rows.Count - 2) //Accounts for the null row
                {
                    OnRecord++;
                    LoadRecord();
                }
                else
                {
                    if (MessageBox.Show("This was the last record! Are you done?", "Last Record", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        OnRecord = 0;
                        SearchRecords(CONN);
                        LoadRecord();
                    }
                }


            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error in AcceptClick: " + Ex.Message + " ... " + Ex.StackTrace);
            }
            finally
            {
                CONN.Close();
            }

            


            
        }

        public void LoadRecord(bool GoingForward = true)
        {
            //Load PDF!
            SqlConnection CONN = new SqlConnection(ConfigurationManager.ConnectionStrings["TESS Dashboard.Properties.Settings.REFTBL4ConnectionString"].ConnectionString);
            CONN.Open();
            try
            {
                //refresh data for the record! MPA 1/29/2020. If the account no longer satisfies the criteria for the search, skip it.
                if (GoingForward && SearchRecords(CONN, (int)DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value))
                {
                    //nothing! Proceed as normal.
                }
                else if (GoingForward)
                {
                    Button_Skip_Click(null, null);
                }

                SQLcommander DocumentCommand = new SQLcommander(CONN);

                List<string> RequestFields = new List<string>();
                RequestFields.Add("ADOC_PATHNAME");

                DocumentCommand.TableName = "tblACCT_DOCUMENT2";

                foreach (string aKey in RequestFields)
                {
                    DocumentCommand.AddRequest(aKey);
                }

                Dictionary<string, Tuple<string, object>> ConditionFields = new Dictionary<string, Tuple<string, object>>();
                ConditionFields.Add("ADOC_ACCT = @ADOC_ACCT", new Tuple<string, object>("@ADOC_ACCT", DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value));
                ConditionFields.Add("ADOC_TYPE LIKE 'app%'", new Tuple<string, object>(null, null));

                foreach (string aKey in ConditionFields.Keys)
                {
                    DocumentCommand.Conditions.Add(aKey);

                    if (ConditionFields[aKey].Item1 != null)
                    {
                        DocumentCommand.Parameters.Add(ConditionFields[aKey].Item1, ConditionFields[aKey].Item2);
                    }
                }

                DocumentCommand.Select();
                object[,] ReturnedRecords = DocumentCommand.SelectResults;

                if (ReturnedRecords.GetUpperBound(0) > 1)
                {
                    MessageBox.Show("There are more than two 'App%' documents for this account!");
                }

                if (ReturnedRecords.GetUpperBound(0) > 0)
                {
                    if (File.Exists(Convert.ToString(ReturnedRecords[0, 0])))
                    {
                        if (!IsFileLocked(Convert.ToString(ReturnedRecords[0, 0])))
                        {
                            //DisplayPDF1.src = "file:////" + Convert.ToString(ReturnedRecords[0, 0]);
                            DisplayPDF1.LoadFile(Convert.ToString(ReturnedRecords[0, 0]));
                            DisplayPDF1.setLayoutMode("TwoColumnLeft");
                            DisplayPDF1.gotoFirstPage();
                            DisplayPDF1.gotoNextPage();

                        }
                        else
                        {
                            MessageBox.Show("File path is locked! " + ReturnedRecords[0, 0].ToString());
                        }

                    }
                    else
                    {
                        MessageBox.Show("File path not found! " + ReturnedRecords[0, 0].ToString());
                    }

                    if (File.Exists(Convert.ToString(ReturnedRecords[1, 0])))
                    {
                        if (!IsFileLocked(Convert.ToString(ReturnedRecords[1, 0])))
                        {

                            DisplayPDF2.LoadFile(Convert.ToString(ReturnedRecords[1, 0]));
                            DisplayPDF2.setLayoutMode("TwoColumnLeft");
                            DisplayPDF2.gotoFirstPage();
                            DisplayPDF2.gotoNextPage();
                            DisplayPDF2.gotoNextPage();
                        }
                        else
                        {
                            MessageBox.Show("File path is locked! " + ReturnedRecords[1, 0].ToString());
                        }

                    }
                    else
                    {
                        MessageBox.Show("File path not found! " + ReturnedRecords[1, 0].ToString());
                    }
                }
                else if (ReturnedRecords.GetUpperBound(0) > -1)
                {

                    if (File.Exists(Convert.ToString(ReturnedRecords[0, 0])))
                    {
                        if (!IsFileLocked(Convert.ToString(ReturnedRecords[0, 0])))
                        {
                            //DisplayPDF1.src = "file:////" + Convert.ToString(ReturnedRecords[0, 0]);
                            DisplayPDF1.LoadFile(Convert.ToString(ReturnedRecords[0, 0]));
                            DisplayPDF1.setLayoutMode("TwoColumnLeft");
                            DisplayPDF1.gotoFirstPage();
                            DisplayPDF1.gotoNextPage();

                            DisplayPDF2.LoadFile(Convert.ToString(ReturnedRecords[0, 0]));
                            DisplayPDF2.setLayoutMode("TwoColumnLeft");
                            DisplayPDF2.gotoFirstPage();
                            DisplayPDF2.gotoNextPage();
                            DisplayPDF2.gotoNextPage();
                        }
                        else
                        {
                            MessageBox.Show("File path is locked! " + ReturnedRecords[0, 0].ToString());
                        }

                    }
                    else
                    {
                        MessageBox.Show("File path not found! " + ReturnedRecords[0, 0].ToString());
                    }

                }
                else
                {
                    DisplayPDF1.LoadFile("sdkjffdjsk");
                    DisplayPDF2.LoadFile("sdlkjfdsjkl");
                    MessageBox.Show("No 'App%' documents were found for this account!");
                }

                //DisplayPDF3.LoadFile(Convert.ToString(ReturnedRecords[0, 0]));
                //DisplayPDF3.gotoFirstPage();
                //DisplayPDF3.gotoNextPage();
                //DisplayPDF3.gotoNextPage();
                //DisplayPDF4.LoadFile(Convert.ToString(ReturnedRecords[0, 0]));
                //DisplayPDF4.gotoFirstPage();
                //DisplayPDF4.gotoNextPage();
                //DisplayPDF4.gotoNextPage();
                //DisplayPDF4.gotoNextPage();

                //The If Null (??) Statements are probably needless considering nulls are already guarded against from being in the grid in the first place


                lbl_ACCT.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value ?? "");
                CheckBox_ACCT_CANCEL.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CANCEL"].Index].Value ?? false);

                CheckBox_ACCT_EX_DP_GOAL.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_EX_DP_GOAL"].Index].Value ?? false);
                CheckBox_ACCT_EX_HEX_GOAL.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_EX_HEX_GOAL"].Index].Value ?? false);
                CheckBox_ACCT_EX_O65_GOAL.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_EX_O65_GOAL"].Index].Value ?? false);

                CheckBox_ACLT_OWNER1_SIGNED.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_OWNER1_SIGNED"].Index].Value ?? false);
                CheckBox_ACLT_OWNER2_SIGNED.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_OWNER2_SIGNED"].Index].Value ?? false);

                CheckBox_ACLT_EX_DP_APPLIED.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_EX_DP_APPLIED"].Index].Value ?? false);
                CheckBox_ACLT_EX_HEX_APPLIED.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_EX_HEX_APPLIED"].Index].Value ?? false);
                CheckBox_ACLT_EX_O65_APPLIED.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_EX_O65_APPLIED"].Index].Value ?? false);

                lbl_ACCT_ALIAS.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_ALIAS"].Index].Value ?? "");

                string strGoal = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_EX_GOAL_PCT"].Index].Value);
                strGoal = strGoal.Length > 0 ? string.Format("{0:#.00}", Convert.ToDecimal(strGoal) * 100) : "";
                strGoal = strGoal.PadRight(strGoal.Length + 1, '&').Contains(".00&") ? strGoal.Replace(".00", "") : strGoal;
                lbl_ACCT_EX_GOAL_PCT.Text = strGoal.Length > 0 ? strGoal + " %" : "";

                string strRefund = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_REFUND"].Index].Value);
                lbl_ACCT_REFUND.Text = strRefund.Length > 0 ? string.Format("{0:#.00}", Convert.ToDecimal(strRefund)) : "";

                lbl_ACCT_TAX_YEAR.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_TAX_YEAR"].Index].Value ?? "");

                string strDOB1 = ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_DOB1"].Index].Value)).ToShortDateString();
                lbl_ACLT_DOB1.Text = strDOB1 == "1/1/0001" ? "" : strDOB1;
                lbl_lbl_DOB1.Visible = !(strDOB1 == "1/1/0001");
                string strDOB2 = ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_DOB2"].Index].Value)).ToShortDateString();
                lbl_ACLT_DOB2.Text = strDOB2 == "1/1/0001" ? "" : strDOB2;
                lbl_lbl_DOB2.Visible = !(strDOB2 == "1/1/0001");

                string strFee = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_FEE_PCT"].Index].Value);
                strFee = strFee.Length > 0 ? string.Format("{0:#.00}", Convert.ToDecimal(strFee) * 100) : "";
                strFee = strFee.PadRight(strFee.Length + 1, '&').Contains(".00&") ? strFee.Replace(".00", "") : strFee;
                lbl_ACLT_FEE_PCT.Text = strFee.Length > 0 ? strFee : "";

                string strSigned = ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_SIGNED"].Index].Value)).ToShortDateString();
                lbl_ACLT_SIGNED.Text = strSigned == "1/1/0001" ? "" : strSigned;

                lbl_ACLT_STATUS.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_STATUS"].Index].Value ?? "");
                lbl_ACLT_MISC_NOTES.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_MISC_NOTES"].Index].Value ?? "");

                lbl_FullName1.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME1_TITLE"].Index].Value ?? "")
                    + " " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME1_FIRST"].Index].Value ?? "")
                    + " " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME1_MIDDLE"].Index].Value ?? "")
                    + " " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME1_LAST"].Index].Value ?? "")
                    + " " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME1_SUFFIX"].Index].Value ?? "");

                lbl_FullName2.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME2_TITLE"].Index].Value ?? "")
                    + " " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME2_FIRST"].Index].Value ?? "")
                    + " " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME2_MIDDLE"].Index].Value ?? "")
                    + " " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME2_LAST"].Index].Value ?? "")
                    + " " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_CONTACT_NAME2_SUFFIX"].Index].Value ?? "");

                lbl_ASLS_SALESMAN.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ASLS_SALESMAN"].Index].Value ?? "");
                lbl_ASLS_STATUS.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ASLS_STATUS"].Index].Value ?? "");

                /* MPA 1/29/2020 if we end up deleting non eligible records, remove this! -- Removed, but keeping the code in case we change our minds
                if (CheckBox_ASLS_STATUS.Checked || (!CheckBox_ASLS_STATUS.Checked && !CheckBox_ASLS_STATUS_TESS.Checked && ((System.Data.DataRowView)ComboBox_ASLS_STATUS.SelectedValue).Row.ItemArray[0].ToString().Replace(" ", "") == ""))
                {
                    lbl_ASLS_STATUS.BackColor = lbl_ASLS_STATUS.Text.ToUpper().Contains("LEAD COMPLETED") ? DefaultBackColor : Color.Yellow;
                }
                else if (CheckBox_ASLS_STATUS_TESS.Checked)
                {
                    lbl_ASLS_STATUS.BackColor = lbl_ASLS_STATUS.Text == "Rep Claims Complete - TESS" ? DefaultBackColor : Color.Yellow;
                }
                else if (!(((System.Data.DataRowView)ComboBox_ASLS_STATUS.SelectedValue).Row.ItemArray[0].ToString() == "All"))
                {
                    lbl_ASLS_STATUS.BackColor = lbl_ASLS_STATUS.Text == ((System.Data.DataRowView)ComboBox_ASLS_STATUS.SelectedValue).Row.ItemArray[0].ToString() ? DefaultBackColor : Color.Yellow;
                }
                */

                lbl_PRCL_PID.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["PRCL_PID"].Index].Value ?? "");
                lbl_PRCL_SITUS_ADDR.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["PRCL_SITUS_ADDR"].Index].Value ?? "");
                lbl_PRCL_SITUS_CITY.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["PRCL_SITUS_CITY"].Index].Value ?? "");
                lbl_PRCL_SITUS_STATE.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["PRCL_SITUS_STATE"].Index].Value ?? "");
                lbl_PRCL_SITUS_ZIP.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["PRCL_SITUS_ZIP"].Index].Value ?? "");

                string strDeed = ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_DEED_DATE"].Index].Value)).ToShortDateString();
                lbl_POWN_DEED_DATE.Text = strDeed == "1/1/0001" ? "" : strDeed;
                lbl_POWN_MAIL_ADDR1.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_MAIL_ADDR1"].Index].Value ?? "");
                lbl_POWN_MAIL_ADDR2.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_MAIL_ADDR2"].Index].Value ?? "");
                lbl_POWN_MAIL_CITY.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_MAIL_CITY"].Index].Value ?? "");
                lbl_POWN_MAIL_STATE.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_MAIL_STATE"].Index].Value ?? "");
                lbl_POWN_MAIL_ZIP.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_MAIL_ZIP"].Index].Value ?? "");
                lbl_POWN_OWNER_NAME1.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_OWNER_NAME1"].Index].Value ?? "");
                lbl_POWN_OWNER_NAME2.Text = Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_OWNER_NAME2"].Index].Value ?? "");

                CheckBox_ACLT_FILE_CURR.Text = Convert.ToString(Convert.ToInt32(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_TAX_YEAR"].Index].Value) - 0);
                CheckBox_ACLT_FILE_LAST.Text = Convert.ToString(Convert.ToInt32(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_TAX_YEAR"].Index].Value) - 1);
                CheckBox_ACLT_FILE_NEXT.Text = Convert.ToString(Convert.ToInt32(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_TAX_YEAR"].Index].Value) + 1);
                CheckBox_ACLT_FILE_PREV.Text = Convert.ToString(Convert.ToInt32(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT_TAX_YEAR"].Index].Value) - 2);

                CheckBox_ACLT_FILE_CURR.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_FILE_CURR"].Index].Value ?? false);
                CheckBox_ACLT_FILE_LAST.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_FILE_LAST"].Index].Value ?? false);
                CheckBox_ACLT_FILE_NEXT.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_FILE_NEXT"].Index].Value ?? false);
                CheckBox_ACLT_FILE_PREV.Checked = (bool)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_FILE_PREV"].Index].Value ?? false);

                //Conditionals

                CheckBox_ACLT_FILE_PREV.BackColor = (!CheckBox_ACLT_FILE_PREV.Checked && (
                    ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_DEED_DATE"].Index].Value)).Year < 2017
                    || ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_DOB1"].Index].Value)).Year <= 1953
                    || ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_DOB2"].Index].Value)).Year <= 1953)
                    ) ? Color.Yellow : CheckBox_ACLT_FILE_PREV.BackColor = DefaultBackColor;
                CheckBox_ACCT_CANCEL.BackColor = CheckBox_ACCT_CANCEL.Checked ? Color.Yellow : DefaultBackColor;
                //CheckBox_ACLT_EX_O65_APPLIED.BackColor = CheckBox_ACLT_EX_O65_APPLIED.Checked ? Color.Yellow : DefaultBackColor;
                lbl_POWN_DEED_DATE.BackColor = ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["POWN_DEED_DATE"].Index].Value)).Year >= 2018 && (
                    ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_DOB1"].Index].Value)).Year <= 1953
                    || ((DateTime)(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACLT_DOB2"].Index].Value)).Year <= 1953
                    ) ? Color.Yellow : DefaultBackColor;

                lbl_ACLT_STATUS.BackColor = lbl_ACLT_STATUS.Text.Length > 0 ? Color.Yellow : DefaultBackColor;

            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error in LoadRecord: " + Ex.Message + " ... " + Ex.StackTrace);
            }
            finally
            {
                CONN.Close();
            }

            lblRecordProgress.Text = "Record " + Convert.ToString(OnRecord + 1) + " of " + Convert.ToString(DataGridView_Sales_Records.Rows.Count - 1); //Accounts for the null row
        }

        private bool IsFileLocked(string FilePath)
        {
            FileInfo file = new FileInfo(FilePath);
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }

        private void DisplayPDF2_Enter(object sender, EventArgs e)
        {

        }

        private void DisplayPDF4_Enter(object sender, EventArgs e)
        {

        }

        private void DisplayPDF3_Enter(object sender, EventArgs e)
        {

        }

        private void lbl_POWN_MAIL_CITY_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void lbl_POWN_EX_VET_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void lbl_POWN_MAIL_STATE_Click(object sender, EventArgs e)
        {

        }

        private void lbl_POWN_MAIL_CITY_Click_1(object sender, EventArgs e)
        {

        }

        private void CheckBox_ACLT_OWNER2_SIGNED_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void Button_Back_Click(object sender, EventArgs e)
        {
            if (OnRecord > 0)
            {
                OnRecord--;
                LoadRecord(false);
            }
            else
            {
                MessageBox.Show("Cannot go back past the first record!");
            }

        }

        private void Button_Needs_Review_Click(object sender, EventArgs e)
        {
            if (ComboBox_NeedsReview.Visible)
            {
                SqlConnection CONN = new SqlConnection(ConfigurationManager.ConnectionStrings["TESS Dashboard.Properties.Settings.REFTBL4ConnectionString"].ConnectionString);
                CONN.Open();
                try
                {
                    SQLcommander UnApproveCommand = new SQLcommander(CONN);
                    UnApproveCommand.TableName = "tblACCT_CLIENT";
                    UnApproveCommand.AddRequest("ACLT_APP_APPROVED = Null");
                    UnApproveCommand.AddCondition("ACLT_ACCT = " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value));
                    UnApproveCommand.Update();
                    
                    SQLcommander SalesCommand = new SQLcommander(CONN);
                    SalesCommand.TableName = "tblACCT_SALES";
                    SalesCommand.AddRequest("ASLS_STATUS = 'Lead Returned to Rep'");
                    SalesCommand.AddCondition("ASLS_ACCT = " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value));
                    SalesCommand.Update();

                    SQLcommander ClientCommand = new SQLcommander(CONN);
                    ClientCommand.TableName = "tblACCT_CLIENT";
                    ClientCommand.AddRequest("ACLT_STATUS = '" + ((System.Data.DataRowView)ComboBox_NeedsReview.SelectedValue).Row.ItemArray[0] + "'");
                    ClientCommand.AddCondition("ACLT_ACCT = " + Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value));
                    ClientCommand.Update();

                    InsertCallRecord(CONN, Convert.ToInt32(Convert.ToString(DataGridView_Sales_Records.Rows[OnRecord].Cells[DataGridView_Sales_Records.Columns["ACCT"].Index].Value)), "TESS forms in pending");

                    if (OnRecord < DataGridView_Sales_Records.Rows.Count - 2) //Accounts for the null row
                    {
                        OnRecord++;
                        LoadRecord();
                    }
                    else
                    {
                        if (MessageBox.Show("This was the last record! Go back to the start?", "Last Record", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            OnRecord = 0;
                            SearchRecords(CONN);
                            LoadRecord();
                        }
                    }
                }
                catch (Exception Ex)
                {
                    MessageBox.Show("Error in NeedsReviewClick: " + Ex.Message + " ... " + Ex.StackTrace);
                }
                finally
                {
                    CONN.Close();
                }

                Button_Needs_Review.Text = "Needs Review";
                ComboBox_NeedsReview.SelectedIndex = 0;
                ComboBox_NeedsReview.Visible = false;
                

            }
            else
            {
                Button_Needs_Review.Text = "Select a reason:";
                ComboBox_NeedsReview.Visible = true;
            }

        }

        private void Button_Resize_Click(object sender, EventArgs e)
        {
            if (Button_Resize.Text == "Resize Up")
            {
                
                DisplayPDF1.Height = DisplayPDF1.Height * 2;
                DisplayPDF2.Height = DisplayPDF2.Height * 2;
                
                DisplayPDF2.Location = new Point(DisplayPDF2.Location.X + DisplayPDF2.Width, DisplayPDF2.Location.Y);

                DisplayPDF1.Width = DisplayPDF1.Width * 2;
                DisplayPDF2.Width = DisplayPDF2.Width * 2;
                
                Button_Resize.Text = "Resize Down";
            }
            else
            {
                DisplayPDF1.Width = DisplayPDF1.Width / 2;
                DisplayPDF2.Width = DisplayPDF2.Width / 2;

                DisplayPDF2.Location = new Point(DisplayPDF2.Location.X - DisplayPDF2.Width, DisplayPDF2.Location.Y);
                
                DisplayPDF1.Height = DisplayPDF1.Height / 2;
                DisplayPDF2.Height = DisplayPDF2.Height / 2;
                

                Button_Resize.Text = "Resize Up";
            }
        }

        private void Button_Skip_Click(object sender, EventArgs e)
        {
            

            SqlConnection CONN = new SqlConnection(ConfigurationManager.ConnectionStrings["TESS Dashboard.Properties.Settings.REFTBL4ConnectionString"].ConnectionString);
            CONN.Open();

            try
            {
                if (OnRecord < DataGridView_Sales_Records.Rows.Count - 2) //Accounts for the null row
                {
                    OnRecord++;
                    LoadRecord();
                }
                else
                {
                    if (MessageBox.Show("This was the last record! Are you done?", "Last Record", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        OnRecord = 0;
                        SearchRecords(CONN);
                        LoadRecord();
                    }
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error in SkipClick: " + Ex.Message + " ... " + Ex.StackTrace);
            }
            finally
            {
                CONN.Close();
            }

        }

        private void lbl_PRCL_SITUS_ZIP_Click(object sender, EventArgs e)
        {

        }

        private void lbl_ACCT_ALIAS_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(lbl_ACCT_ALIAS.Text);
        }

        private void lbl_ACCT_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(lbl_ACCT.Text);
        }

        private void lbl_PRCL_PID_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(lbl_PRCL_PID.Text);
        }

        private void ComboBox_ASLS_STATUS_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void CheckBox_ASLS_STATUS_CheckedChanged(object sender, EventArgs e)
        {
            ASLS_STATUS_visibilityupdate();
        }

        private void ComboBox_County_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void CheckBox_ASLS_STATUS_TESS_CheckedChanged(object sender, EventArgs e)
        {
            ASLS_STATUS_visibilityupdate(true);
        }

        private void ASLS_STATUS_visibilityupdate(bool IsTess = false)
        {
            CheckBox_ASLS_STATUS.Checked = IsTess ? !CheckBox_ASLS_STATUS_TESS.Checked : CheckBox_ASLS_STATUS.Checked;
            CheckBox_ASLS_STATUS_TESS.Checked = IsTess ? CheckBox_ASLS_STATUS_TESS.Checked : !CheckBox_ASLS_STATUS.Checked;
            //ComboBox_ASLS_STATUS.Visible = !(CheckBox_ASLS_STATUS.Checked || CheckBox_ASLS_STATUS_TESS.Checked); //MPA 12/18/2019 disabling dropdown for now
            Button_Accept.Text = CheckBox_ASLS_STATUS.Checked ? "Approve" : Button_Accept.Text;
            Button_Accept.Text = CheckBox_ASLS_STATUS_TESS.Checked ? "Accept" : Button_Accept.Text;
        }

        private bool InsertCallRecord(SqlConnection conn, int acct, string notes = "")
        {
            string sql = "INSERT INTO tblACCT_CALL (ACLL_ACCT, ACLL_TYPE, ACLL_INITIALS, ACLL_DATE, ACLL_TIME, ACLL_ORG, ACLL_NOTES) VALUES (@ACCT, @TYPE, @INI, @DATE, @TIME, @ORG, @NOTES)";
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@ACCT", acct);
            cmd.Parameters.AddWithValue("@TYPE", "T");
            cmd.Parameters.AddWithValue("@INI", GetInitials());
            cmd.Parameters.AddWithValue("@DATE", DateTime.Now);
            cmd.Parameters.AddWithValue("@TIME", DateTime.Now.ToString("hh:mm tt"));
            cmd.Parameters.AddWithValue("@ORG", "RAC");
            cmd.Parameters.AddWithValue("@NOTES", notes);
            return cmd.ExecuteNonQuery() > 0;
        }

        private string GetInitials()
        {
            PrincipalContext pc = new PrincipalContext(ContextType.Domain, "corp.kwardcpa.com");
            string username = "";
            username = Environment.UserName;

            UserPrincipal user = UserPrincipal.FindByIdentity(pc, username);
            string distinguished = user.DistinguishedName;

            IADsUser Usr = (IADsUser)Microsoft.VisualBasic.Interaction.GetObject("LDAP://corp.kwardcpa.com/" + distinguished);
            return Usr.Get("initials");
        }

        private void Main_Window_FormClosed(object sender, FormClosedEventArgs e)
        {
            DisplayPDF1.Dispose();
            DisplayPDF2.Dispose();
        }

        private void DisplayPDF1_Enter(object sender, EventArgs e)
        {

        }

        private void Button_PDF1_Open_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(DisplayPDF1.src))
            {
                using (Process myProcess = new Process())
                {
                    myProcess.StartInfo.UseShellExecute = true;
                    myProcess.StartInfo.FileName = @"C:\Program Files (x86)\PDF Pro 10\PDFEditor.exe";
                    myProcess.StartInfo.Arguments = DisplayPDF1.src;
                    myProcess.Start();
                }
            }
            else
            {
                MessageBox.Show("PDF1 does not point to a findable file path!");
            }

        }

        private void Button_PDF2_Open_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(DisplayPDF2.src))
            {
                using (Process myProcess = new Process())
                {
                    myProcess.StartInfo.UseShellExecute = true;
                    myProcess.StartInfo.FileName = @"C:\Program Files (x86)\PDF Pro 10\PDFEditor.exe";
                    myProcess.StartInfo.Arguments = DisplayPDF2.src;
                    myProcess.Start();
                }
            }
            else
            {
                MessageBox.Show("PDF2 does not point to a findable file path!");
            }
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void lbl_ASLS_STATUS_Click(object sender, EventArgs e)
        {

        }
    }
}

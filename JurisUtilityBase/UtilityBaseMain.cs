using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        private string error = "";

        private string bCodesForTest = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
            //            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
            //            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }
            UpdateBanks();
        }

   
            private void UpdateBanks()
            {
                cbBank.ClearItems();
                cbNew.ClearItems();
                string SqlBank = "select left(bnkcode + '          ',10) + bnkdesc as Bank from BankAccount order by bnkcode";

                DataSet sb = _jurisUtility.RecordsetFromSQL(SqlBank);
                string BCode;


                if (sb.Tables[0].Rows.Count == 0)
                    cbBank.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table in sb.Tables)
                    {

                        foreach (DataRow dr in table.Rows)
                        {
                            BCode = dr["Bank"].ToString();
                            cbBank.Items.Add(BCode);
                        }
                    }
                }

                string SqlNew = "select Bank from (select left('****' + '          ',10) + 'CREATE NEW' as bank union all select left(bnkcode + '          ',10) + bnkdesc as Bank from BankAccount) BK group by bank order by bank";

                DataSet sn = _jurisUtility.RecordsetFromSQL(SqlNew);
                string NCode;


                if (sn.Tables[0].Rows.Count == 0)
                    cbNew.SelectedIndex = 0;
                else
                {
                    foreach (DataTable table2 in sn.Tables)
                    {

                        foreach (DataRow dr2 in table2.Rows)
                        {
                            NCode = dr2["Bank"].ToString();
                            cbNew.Items.Add(NCode);
                        }
                    }
                }
            }
   
        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

            string OBank = cbBank.SelectedItem.ToString();
            string OldBank = OBank.Substring(0, 4);
            OldBank = OldBank.TrimEnd(' ');
            string NBank = cbNew.SelectedItem.ToString();
            string NewBank = NBank.Substring(0, 4);
            NewBank = NewBank.TrimEnd(' ');


                if (NewBank.ToString() == "****")
                {
                    NewBank = txtBnkCode.Text.ToString();
                    NewBank = NewBank.TrimEnd(' ');
                    string NewDesc = txtBankDesc.Text.ToString();

                    string Sql = "select bnkcode from bankaccount where bnkcode='" + NewBank.ToString() + "'";

                    DataSet sn = _jurisUtility.RecordsetFromSQL(Sql);

                    if (sn.Tables[0].Rows.Count == 0)
                    {
                        DialogResult rs = MessageBox.Show("Bank account " + OldBank.ToString() + " will be moved to " + NewBank.ToString() + " - " + NewDesc.ToString() + ".  Do you wish to continue?", "Warning",
                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                        if (rs == DialogResult.Yes)
                        { RenameBank(NewBank, OldBank, NewDesc); }
                        else
                        {
                            labelCurrentStatus.Text = "Operation Cancelled";
                            toolStripStatusLabel.Text = "Operation Cancelled";
                        }
                    }
                    else
                    {
                        DialogResult rs2 = MessageBox.Show("Bank account " + NewBank.ToString() + " exists within the current database. Please enter a new bank code.", "Alert",
                      MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                    }

                }

                else
                {
                                string sqlquery = "select BnkAcctType from BankAccount where bnkcode = '" + NewBank + "' or bnkcode = '" + OldBank + "'";
                        DataSet ds4 = _jurisUtility.RecordsetFromSQL(sqlquery);
                        if (ds4.Tables[0].Rows[0][0].ToString().Equals(ds4.Tables[0].Rows[1][0].ToString()))
                        {
                    DialogResult rs = MessageBox.Show("Bank account " + OldBank.ToString() + " will be merged with existing bank " + NewBank.ToString() + " and " + OldBank.ToString() + " will be removed from the database.  Do you wish to continue?", "Warning",
                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    if (rs == DialogResult.Yes)
                    {
                        MergeBank(NewBank, OldBank);
                    }
                    else if (rs == DialogResult.No)
                    {
                        //code for No
                    }
                    else if (rs == DialogResult.Cancel)
                    {
                        //code for Cancel
                    }
            }
            else
                MessageBox.Show("You can only select the same type of bank accounts." + "\r\n" + "For example, you cannot replace a General Bank with a Trust Bank", "Selection Issue", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }


        }

  



        private void MergeBank(string nb, string ob)
        {
            UpdateStatus("Merging Banks...", 1, 20);
            toolStripStatusLabel.Text = "Merging Banks...";
            Cursor.Current = Cursors.WaitCursor;
            statusStrip.Refresh();

            if (disableConstraints())
            {

                string sql = "update ARMatTrust set ARMTBank='" + nb + "' where armtbank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 2, 20);
                sql = "update CBCheck set CBCBank='" + nb + "' where CBCBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 3, 20);
                sql = "update CheckRegister set ckregcleared='Y’, CkRegBank='" + nb + "' where CkRegBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 4, 20);
                sql = "update CheckRegister_Log set CkRegBank='" + nb + "' where CkRegBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 5, 20);
                sql = "update CRARAlloc set CRABank='" + nb + "' where CRABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 6, 20);
                sql = "update CRARAlloc_Log set CRABank='" + nb + "' where CRABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 7, 20);
                sql = "update CRNonCliAlloc set CRNBankCode='" + nb + "' where CRNBankCode='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 8, 20);
                sql = "update CRTrustAlloc set CRTBank='" + nb + "' where CRTBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 9, 20);
                sql = "update OfficeCode set OfcBankCode='" + nb + "' where OfcBankCode='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 10, 20);
                sql = "update PrebillExpenseTrustApplied set PBETABank='" + nb + "' where PBETABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 11, 20);
                sql = "update PrebillFeeTrustApplied set PBFTABank='" + nb + "' where PBFTABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 12, 20);
                sql = "update PrebillMatterTrustApplied set PBMTABank='" + nb + "' where PBMTABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 13, 20);
                sql = "update TrAdjBatchDetail set TABDBank='" + nb + "' where TABDBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 14, 20);

                UpdateStatus("Merging Banks...", 15, 20);
                sql = "update TrustLedger set TLBank='" + nb + "' where TLBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 16, 20);

                UpdateStatus("Merging Banks...", 17, 20);
                sql = "update VchTemplate set VTVchBank='" + nb + "' where VTVchBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 18, 20);
                sql = "update Voucher set VchBank='" + nb + "' where VchBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 19, 20);
                sql = "update VoucherBatchDetail set VBDBank='" + nb + "' where VBDBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Merging Banks...", 20, 20);
                sql = "update VoucherPayment set VPBank= '" + nb + "' where VPBank = '" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Removing Old Bank...", 1, 3);

                string USql = @"select '" + nb + @"' as bnk, tspmatter as Mat, tspprdyear as PrdYear, tspprdnbr as Prd, sum(tspdeposits) as Dep, sum(tsppayments) as Pay, sum(tspadjustments) as Adj 
                 Into #Tsp
                from trustsumbyprd
                where tspbank='" + nb + "' or tspbank='" + ob + "' group by tspmatter, tspprdnbr, tspprdyear";
                _jurisUtility.ExecuteNonQueryCommand(0, USql);

                sql = "delete from trustsumbyprd  where tspbank='" + nb + "' or tspbank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);

                sql = "insert into trustsumbyprd(tspbank,tspmatter, tspprdyear, tspprdnbr, tspdeposits, tsppayments, tspadjustments) select bnk,mat, prdyear, prd, dep,pay, adj from #Tsp";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);

                sql = "Drop table #Tsp";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);

                sql = "update TrustAccount set TABank='" + nb + "' where TABank='" + ob + "' and tamatter not in (select tamatter from trustaccount where tabank='" + nb + "')";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);

                sql = @"update TrustAccount set tabalance=tbal, TADateLastActy=lastact from  (select '" + nb + @"' as bank, tlmatter, max(tldate) as lastact, sum(tlamount) as tbal 
                from trustledger where tlbank in ('" + ob + "','" + nb + "') group by TlMatter) TA  where  tabank=bank  and tamatter=tlmatter";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "delete from trustaccount where tabank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "delete from bkacctglacct where bgabkacct= '" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Removing Old Bank...", 2, 3);

                sql = "select '" + nb + "' as brhbank, brhstmtdate, sum(brhstmtopenbal) as OpenBal, sum(brhstmtdepositcount) as DepositCount, sum(brhstmtdepositamount) as DepAmt, "
                    + " sum(brhstmtcheckcount) as CheckCount, sum(brhstmtcheckamount) as CheckAmt, sum(brhstmtendbal) as EndBal,"
                    + " max( brhbooklaststmtdate) as LastStmtDate, "
                    + " sum(brhbooklaststmtbal) as LastBal, sum(brhbookdepositcount) as BookDepCount, sum(brhbookdepositamount) as BookDepAmount,"
                    + " sum(brhbookcheckcount) as BookCheckCount, sum(brhbookcheckamount) as BookCheckAmt, sum(brhbookclearedbal) as BookCleared, max(brhrecorded) as Recorded, max(brhrecordeddate) as RecordedDate,"
                    + " max(brhlastckregbatchje) as LastBatchJE, row_number() over (order by brhstmtdate) as RowID"
                  //  + " into #BRH"
                    + " from bankreconhistory where brhrecorded='Y' and   (brhbank='" + nb + "' or brhbank='" + ob + "')"
                    + " group by  brhstmtdate order by brhstmtdate asc";
                DataSet ds22 = _jurisUtility.RecordsetFromSQL(sql);

                ds22.Tables[0].DefaultView.Sort = "RowID asc";
                DataTable dt = ds22.Tables[0].DefaultView.ToTable();

                sql = "delete from bankreconhistory where  brhbank='" + nb + "' or brhbank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);

                foreach (DataRow row in dt.Rows)
                {

                    sql = "Insert into bankreconhistory(brhbank, brhstmtdate, brhstmtopenbal,brhstmtdepositcount,brhstmtdepositamount,brhstmtcheckcount,brhstmtcheckamount,brhstmtendbal,"
                               + " brhbooklaststmtdate,brhbooklaststmtbal,brhbookdepositcount,brhbookdepositamount,brhbookcheckcount,brhbookcheckamount,brhbookclearedbal,brhrecorded,brhrecordeddate,"
                               + " brhlastckregbatchje) values ('"
                               + row["brhbank"].ToString().Trim() + "', '" + Convert.ToDateTime(row["brhstmtdate"].ToString()).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture) + "' , " + row["OpenBal"].ToString().Trim() +
                               ", " + row["DepositCount"].ToString().Trim() + ", " + row["DepAmt"].ToString().Trim() + " , " +
                               row["CheckCount"].ToString().Trim() + " , " + row["CheckAmt"].ToString().Trim() + " , " + row["EndBal"].ToString().Trim() + " , '" + Convert.ToDateTime(row["LastStmtDate"].ToString().Trim()).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture) +
                               "' , " + row["LastBal"].ToString().Trim() + " , " + row["BookDepCount"].ToString().Trim() + " , " + row["BookDepAmount"].ToString().Trim() +
                               " , " + row["BookCheckCount"].ToString().Trim() + " , " + row["BookCheckAmt"].ToString().Trim() + " , " + row["BookCleared"].ToString().Trim() +
                               " , '" + row["Recorded"].ToString().Trim() + "' , '" + Convert.ToDateTime(row["RecordedDate"].ToString().Trim()).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture) + "' , " + row["LastBatchJE"].ToString().Trim() + " )"; 
                    _jurisUtility.ExecuteNonQueryCommand(0, sql);
                }

                //sql = "drop table #BRH";
               // _jurisUtility.ExecuteNonQueryCommand(0, sql);

                UpdateStatus("Removing Old Bank...", 3, 3);


                sql = "delete from bankaccount where bnkcode= '" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "delete from documenttree where DTKeyT= '" + ob + "' and DTDocClass = 6400 and DTDocType = 'R' and DTParentID = 10";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);

                enableConstraints();
                UpdateStatus("Complete...", 3, 3);
                toolStripStatusLabel.Text = "Complete...";
                Cursor.Current = Cursors.Default;
                statusStrip.Refresh();

                UpdateBanks();
            }
            else
                MessageBox.Show("There was a permission issue in the books. No changes were made. Details: " + "\r\n" + error, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


        private void RenameBank(string nb, string ob, string bdesc)
      {
            string sql = "";
            disableConstraints();

                UpdateStatus("Renaming Bank...", 1, 20);
                toolStripStatusLabel.Text = "Renaming Bank...";
                Cursor.Current = Cursors.WaitCursor;
                statusStrip.Refresh();

                sql = "update ARMatTrust set ARMTBank='" + nb + "' where armtbank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 2, 20);
                sql = "update CBCheck set CBCBank='" + nb + "' where CBCBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 3, 20);
                sql = "update CheckRegister set CkRegBank='" + nb + "' where CkRegBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 4, 20);
                sql = "update CheckRegister_Log set CkRegBank='" + nb + "' where CkRegBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 5, 20);
                sql = "update CRARAlloc set CRABank='" + nb + "' where CRABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 6, 20);
                sql = "update CRARAlloc_Log set CRABank='" + nb + "' where CRABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 7, 20);
                sql = "update CRNonCliAlloc set CRNBankCode='" + nb + "' where CRNBankCode='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 8, 20);
                sql = "update CRTrustAlloc set CRTBank='" + nb + "' where CRTBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 9, 20);
                sql = "update OfficeCode set OfcBankCode='" + nb + "' where OfcBankCode='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 10, 20);
                sql = "update PrebillExpenseTrustApplied set PBETABank='" + nb + "' where PBETABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 11, 20);
                sql = "update PrebillFeeTrustApplied set PBFTABank='" + nb + "' where PBFTABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 12, 20);
                sql = "update PrebillMatterTrustApplied set PBMTABank='" + nb + "' where PBMTABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "update TrAdjBatchDetail set TABDBank='" + nb + "' where TABDBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 13, 20);
                sql = "update TrustLedger set TLBank='" + nb + "' where TLBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "update VchTemplate set VTVchBank='" + nb + "' where VTVchBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 14, 20);
                sql = "update Voucher set VchBank='" + nb + "' where VchBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 15, 20);
                sql = "update VoucherBatchDetail set VBDBank='" + nb + "' where VBDBank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 16, 20);
                sql = "update VoucherPayment set VPBank= '" + nb + "' where VPBank = '" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 17, 20);
                sql = "update TrustSumbyPrd set TSpBank= '" + nb + "' where TspBank = '" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);

                UpdateStatus("Renaming Bank...", 18, 20);
                sql = "update TrustAccount set TABank='" + nb + "' where TABank='" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "update bkacctglacct set bgabkacct='" + nb + "' where bgabkacct= '" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 19, 20);
                sql = "update bankreconhistory set brhbank='" + nb + "'  where brhbank= '" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                UpdateStatus("Renaming Bank...", 20, 20);
                sql = "update  bankaccount set bnkcode='" + nb + "', bnkdesc='" + bdesc + "' where bnkcode= '" + ob + "'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "delete from organizationalunitteammember where parentid not in (select id from organizationalunit)";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "delete from documenttree where DTKeyT= '" + ob + "' and DTDocClass = 6400 and DTDocType = 'R' and DTParentID = 10";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "insert into documenttree ([DTDocID], [DTSystemCreated] ,[DTDocClass],[DTDocType],[DTParentID] ,[DTTitle],[DTKeyL],[DTKeyT]) " +
                            "values ((select max(DTDocID) + 1 from documenttree), 'Y', 6400, 'R', 10, '" + bdesc + "', null, '" + nb + "')";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "update sysparam set SpNbrValue = (select max(DTDocID) from documenttree) where spname = 'LastSysNbrDocTree'";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                enableConstraints();


            UpdateStatus("Complete...", 20, 20);
            toolStripStatusLabel.Text = "Complete...";
            Cursor.Current = Cursors.Default;
            statusStrip.Refresh();

            UpdateBanks();
        }

        private bool disableConstraints()
        {
            try
            {
                string sql = "ALTER TABLE ARMatTrust NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE CBCheck NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE CheckRegister NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE CRARAlloc NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE CRNonCliAlloc NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE CRTrustAlloc NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE OfficeCode NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE PrebillExpenseTrustApplied NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE PrebillFeeTrustApplied NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE PrebillMatterTrustApplied NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE TrAdjBatchDetail NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE TrustLedger NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE VchTemplate NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE Voucher NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE VoucherBatchDetail NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE VoucherPayment NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE trustsumbyprd NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE TrustAccount NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE bkacctglacct NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE bankreconhistory NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE bankaccount NOCHECK CONSTRAINT ALL"; 
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE organizationalunitteammember NOCHECK CONSTRAINT ALL";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE documenttree NOCHECK CONSTRAINT ALL";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);
                sql = "ALTER TABLE sysparam NOCHECK CONSTRAINT ALL";
                _jurisUtility.ExecuteNonQueryCommand(0, sql);

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }


        }

        private void enableConstraints()
        {
            string sql = "ALTER TABLE ARMatTrust CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE CBCheck CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE CheckRegister CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE CRARAlloc CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE CRNonCliAlloc CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE CRTrustAlloc CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE OfficeCode CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE PrebillExpenseTrustApplied CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE PrebillFeeTrustApplied CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE PrebillMatterTrustApplied CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE TrAdjBatchDetail CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE TrustLedger CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE VchTemplate CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE Voucher CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE VoucherBatchDetail CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE VoucherPayment CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE trustsumbyprd CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE TrustAccount CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE bkacctglacct CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE bankreconhistory CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE bankaccount CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE documenttree CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
            sql = "ALTER TABLE sysparam CHECK CONSTRAINT ALL";
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }




        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
           // if (string.IsNullOrEmpty(toAtty) || string.IsNullOrEmpty(fromAtty))
          //      MessageBox.Show("Please select from both Timekeeper drop downs", "Selection Error");
          //  else
          //  {
                //generates output of the report for before and after the change will be made to client
                string SQLTkpr = getReportSQL();

                DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

                ReportDisplay rpds = new ReportDisplay(myRSTkpr);
                rpds.Show();

           // }
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            //if matter and billing timekeeper
            if (true)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentBillingTimekeeper, 'DEF' as NewBillingTimekeeper" +
                        " from matter" +
                        " inner join client on matclinbr=clisysnbr" +
                        " inner join billto on matbillto=billtosysnbr" +
                        " inner join employee on empsysnbr=billtobillingatty" +
                        " where empinitials<>'ABC'";


            //if matter and originating timekeeper
            else if (false)
                reportSQL = "select Clicode, Clireportingname, Matcode, Matreportingname,empinitials as CurrentOriginatingTimekeeper, 'DEF' as NewOriginatingTimekeeper" +
                    " from matter" +
                    " inner join client on matclinbr=clisysnbr" +
                    " inner join matorigatty on matsysnbr=morigmat" +
                    " inner join employee on empsysnbr=morigatty" +
                    " where empinitials<>'ABC'";


            return reportSQL;
        }

        private void cbNew_SelectedIndexChanged(object sender, EventArgs e)
        {
            string NewBank = cbNew.Text.ToString();
            String BnkCode = NewBank.Substring(0, 4);

            if(BnkCode.ToString()== "****")
            { lblBank.Visible = true;
                txtBnkCode.Visible = true;
                txtBankDesc.Visible = true;
            }
            else
            {
                lblBank.Visible = false;
                txtBnkCode.Visible = false;
                txtBankDesc.Visible = false;
            }

        }

        private void cbBank_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

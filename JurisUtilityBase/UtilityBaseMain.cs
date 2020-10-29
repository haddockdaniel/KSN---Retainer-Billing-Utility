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
using System.Data.SqlClient;
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

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        public string billToAttyEmpSys { get; set; }

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

            string sql = "SELECT distinct empname, BillToBillingAtty FROM PreBill " +
                            "inner join billto on billtosysnbr = pbbillto " +
                            "inner join employee on empsysnbr = BillToBillingAtty " +
                            " inner join prebillmatter on pbmprebill = pbsysnbr " +
                            " inner join matter on matsysnbr = pbmmatter " +
                            " where pbstatus <= 2 and matbillagreecode = 'R' and matfltfeeorretainer<>0 and matstatusflag='O' ";
            DataSet emp = _jurisUtility.RecordsetFromSQL(sql);
            if (emp == null || emp.Tables[0].Rows.Count == 0)
            {  MessageBox.Show("There are no prebills to process", "No processing", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                                      }


            
            else
            {
                comboBox1.ValueMember = "BillToBillingAtty";
                comboBox1.DisplayMember = "empname";
                comboBox1.DataSource = emp.Tables[0];
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            toolStripStatusLabel.Text = "Processing Retainer Prebill Allocations....";
            statusStrip.Refresh();

            string SQLSel = @"select cast(Prebill as varchar(20)) as Prebill, cast(matsysnbr as varchar(10)) as MatterSys,cast(billtobillingatty as varchar(20)) as BTkpr, cast(morig1 as varchar(20)) as OTkpr1,
            cast(case when morig2 is null then '' else morig2 end as varchar(20)) as OTkpr2, 
                cast(case when morig3 is null then '' else morig3 end as  varchar(20)) OTkpr3
        ,convert(varchar(30),cast(case when retainertoallocate>0 then (retainertoallocate * OTPct1 * .01)  else 0 end as money),1) as OT1Alloc1
,case when convert(varchar(30),cast(case when retainertoallocate>0 then (retainertoallocate * OTPct2 * .01)  else 0 end as money),1) is null then '0.00' else convert(varchar(30),cast(case when retainertoallocate>0 then (retainertoallocate * OTPct2 * .01)  else 0 end as money),1) end as OT2Alloc2
,case when convert(varchar(30),cast(case when retainertoallocate>0 then (retainertoallocate * OTPct3 * .01)  else 0 end as money),1) is null then '0.00' else convert(varchar(30),cast(case when retainertoallocate>0 then (retainertoallocate * OTPct3 * .01)  else 0 end as money),1) end  as OT3Alloc3,
convert(varchar(30),case when retainertoallocate>0 then (
cast((retainertoallocate * case when otpct1 is null then '0' when OTPct1='' then '0' else otPct1 end * .01) as decimal(12,2)) +
cast((retainertoallocate * case when otpct2 is null then '0' when OTPct2='' then '0' else otpct2 end * .01) as decimal(12,2)) +
cast((retainertoallocate * case when otpct3 is null then '0' when  OTPct3='' then '0' else otpct3 end  * .01) as decimal(12,2))) - retainertoallocate else 0  end,1) as Rmd

            from (select prebill as Prebill,billtobillingatty, matsysnbr, matfltfeeorretainer as retainer, sum(totalfees) as TotalFees,  morig1, morig2, morig3, morig4, 
            sum(cast(case when totalfees is null then 0 else totalfees end  as money)) as Total,
            sum(cast(otherfees as money)) as Other,  sum(cast(otkpr1Fees as money))  as O1Fees, ot1, case when cast(otpct1 as int)=0 then '' else  cast(otpct1 as varchar(10)) end  as OTPct1, 
            sum(cast(otkpr2Fees as money)) as O2Fees, ot2, case when cast(otpct2 as int)=0 then '' else  cast(otpct2 as varchar(10)) end  as OTPct2, 
            sum(cast(otkpr3Fees as money)) as O3Fees, ot3, case when cast(otpct3 as int)=0 then '' else  cast(otpct3 as varchar(10)) end   as OTPct3, 
            sum(cast(otkpr4Fees as money)) as O4Fees, ot4, case when cast(otpct4 as int)=0 then '' else  cast(otpct4 as varchar(10)) end  as OTPct4,
           case when matfltfeeorretainer is null then 0 else matfltfeeorretainer end - sum(cast(case when totalfees is null then 0 else totalfees end as money)) as RetainerToAllocate
            from (select pbmprebill as prebill, utmatter, utamount as TotalFees, case when uttkpr not in (select morigatty from matorigatty where morigmat=matsysnbr) then utamount else null end as OtherFees
            ,case when uttkpr=morig1 then utamount else null end as Otkpr1Fees, case when uttkpr=morig2 then utamount else null end as Otkpr2Fees
            ,case when uttkpr=morig3 then utamount else null end as Otkpr3Fees, case when uttkpr=morig4 then utamount else null end as Otkpr4Fees,morig1, morig2, morig3, morig4, ot1, ot2, ot3, ot4, otpct1, otpct2, otpct3, otpct4, null as Expenses
            from unbilledtime
            inner join prebillmatter on pbmmatter=utmatter
            inner join prebillfeeitem on pbfutbatch=utbatch and pbfutrecnbr=utrecnbr
            inner join matter on pbmmatter=matsysnbr
            inner join (select morigmat, max(case when rnk=1 then morigatty else '' end) as morig1, max(case when rnk=2 then morigatty else null end) as morig2, max(case when rnk=3 then morigatty else null end) as morig3
            , max(case when rnk=4 then morigatty else null end) as morig4, max(case when rnk=1 then empinitials else null end) as OT1, max(case when rnk=2 then empinitials else null end) as OT2
            , max(case when rnk=3 then empinitials else null end) as OT3, max(case when rnk=4 then empinitials else null end) as OT4, max(case when rnk=1 then cast(morigpcnt as int) else null end) as OTpct1
            , max(case when rnk=2 then cast(morigpcnt as int) else null end) as OTpct2, max(case when rnk=3 then cast(morigpcnt as int) else null end) as OTpct3
            , max(case when rnk=4 then cast(morigpcnt as int) else null end) as OTpct4
            from (select morigmat, morigatty, morigpcnt, empinitials, rank() over (partition by morigmat order by case when billtobillingatty=morigatty then 0 else morigatty end) as rnk
            from matorigatty
            inner join employee on empsysnbr=morigatty
            inner join matter on morigmat=matsysnbr
            inner join billto on matbillto=billtosysnbr)MO
            group by morigmat)MOrig on morigmat=matsysnbr
            where matbillagreecode='R' and matfltfeeorretainer<>0 and matstatusflag='O'
   
            ) UT
            inner join matter on matsysnbr=utmatter
            inner join client on matclinbr=clisysnbr
            inner join billto on matbillto=billtosysnbr 
            inner join prebill on prebill=pbsysnbr
            where pbstatus<=2 and matbillagreecode='R' and pbbillto in (select billtosysnbr from billto where billtobillingatty = " + billToAttyEmpSys + ") " +
            " and (cast(billtobillingatty as varchar(20))<>cast(morig1 as varchar(20))  or " +
            " cast(case when morig2 is null then '' else morig2 end as varchar(20))>'')  " +
            " group by prebill, billtobillingatty, matsysnbr, matfltfeeorretainer,  morig1, morig2, morig3, morig4, ot1, ot2, ot3, ot4, otpct1, otpct2, otpct3, otpct4 " +
            " ) AllocTbl order by prebill";

            DataSet PBRS = _jurisUtility.RecordsetFromSQL(SQLSel);

            int counter;

            int rowCount = PBRS.Tables[0].Rows.Count;

            string Prebill = "";
            string MatSys = "";
            string Btkpr = "";
            string Otkpr1 = "";
            string Otkpr2 = "";
            string Otkpr3 = "";
            string Alloc1 = "";
            string Alloc2 = "";
            string Alloc3 = "";
            string Rd = "";

            string SQLCT = "EXEC sp_MSforeachtable @command1='ALTER TABLE ? NOCHECK CONSTRAINT ALL'";
            _jurisUtility.ExecuteNonQueryCommand(0, SQLCT);


            if (rowCount != 0)
            {

                counter = 0;
                foreach (DataRow row in PBRS.Tables[0].Rows)
                {
                    counter = counter + 1;
                    Prebill = row["Prebill"].ToString();
                    toolStripStatusLabel.Text = "Processing Retainer Prebill Allocations for " + Prebill.ToString() + "....";
                    statusStrip.Refresh();
                    MatSys = row["MatterSys"].ToString();
                    Btkpr = row["BTkpr"].ToString();
                    Otkpr1 = row["OTkpr1"].ToString();
                    Otkpr2 = row["OTkpr2"].ToString();
                    Otkpr3 = row["OTkpr3"].ToString();
                    Alloc1 = row["OT1Alloc1"].ToString();
                    Alloc2 = row["OT2Alloc2"].ToString();
                    Alloc3 = row["OT3Alloc3"].ToString();
                    Rd = row["Rmd"].ToString();

                    string SQL99 = @"update prebillfeeitem set pbfamtonbill=utamount from unbilledtime where utbatch=pbfutbatch and utrecnbr=pbfutrecnbr and pbfstatus='0' and pbfprebill= cast('" + Prebill.ToString() + "' as int)  and cast('" + Alloc1.ToString() + "' as money)>0";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL99);

                    string SQL10 = @"update prebillfeeitem set pbfamtonbill=cast('" + Alloc1.ToString() + "' as money), pbftkpronbill=cast('" + Otkpr1.ToString() + "' as int)  where pbftkpronbill = cast('" + Btkpr.ToString() + "' as int) and pbfseqonbill=0  and pbfhrsonbill = 0 and pbfprebill = cast('" + Prebill.ToString() + "' as int) and pbfstatus = 'I' and cast('" + Alloc1.ToString() + "' as money)>0";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL10);

                    string SQL436 = @"insert into timebatchdetail(tbdbatch,tbdrecnbr, tbdrectype, tbdposted, tbddate, tbdprdyear, tbdprdnbr, tbdtkpr, tbdmatter, tbdbudgphase, tbdfeesched, tbdtaskcd, tbdactivitycd,tbdbillableflg,tbdactualhrswrk, tbdhourssource, tbdhourstobill, tbdratesource, tbdrate, tbdamountsource, tbdamount, tbdcode1, tbdcode2, tbdcode3, tbdbillnote, tbdstopwatch, tbdnarrative, tbdid)
                    Select pbbatch, maxrec +1 as Recnbr, 2, 'Y',pbfd, (select prdyear from actngperiod where prdstartdate<=pbfd and prdenddate>=pbfd and prdnbr<>0), 
                    (select prdnbr from actngperiod where prdstartdate<=pbfd and prdenddate>=pbfd and prdnbr<>0),cast('" + Otkpr1.ToString() + @"' as int),cast('" + MatSys.ToString() + @"' as int), 0, matfeesch, null,null, 'Y', 0,'W',0,3, 0, 3, 0, '','','','',0,'', (select max(tbdid) + 1 from timebatchdetail)
                    from  prebillmatter
                    inner join matter on pbmmatter=matsysnbr
                    left outer join tkprrate on tkrfeesch=matfeesch and tkremp=cast('" + Otkpr1.ToString() + @"' as int) 
                    left outer join perstyprate on ptrfeesch=matfeesch and ptrprstyp=(select empprstyp from employee where empsysnbr=cast('" + Otkpr1.ToString() + @"' as int)) , 
(select tbdbatch as pbbatch, max(tbdrecnbr) as maxrec, (select max(pbfdateonbill) from prebillfeeitem) as pbfd
 from timebatchdetail     where tbdbatch=(select max(pbfutbatch) from prebillfeeitem )  group by tbdbatch) PBatch 
where  cast('" + Otkpr1.ToString() + @"' as int) is not null and
cast('" + Otkpr1.ToString() + "' as int) > 0 and cast('" + Alloc1.ToString() + @"' as money) <> 0
  and cast('" + Alloc1.ToString() + "' as money) is not null and pbmmatter = cast('" + MatSys.ToString() + @"' as int)
  and  pbmprebill = cast('" + Prebill.ToString() + "' as int) and cast('" + Prebill.ToString() + "' as int) not in (select pbfprebill from prebillfeeitem where pbfstatus='I' and pbftkpronbill=cast('" + Otkpr1.ToString() + @"' as int)) ";

                    _jurisUtility.ExecuteNonQueryCommand(0, SQL436);


                    string SQL46 = @"insert into timebatchdetail(tbdbatch,tbdrecnbr, tbdrectype, tbdposted, tbddate, tbdprdyear, tbdprdnbr, tbdtkpr, tbdmatter, tbdbudgphase, tbdfeesched, tbdtaskcd, tbdactivitycd,tbdbillableflg,tbdactualhrswrk, tbdhourssource, tbdhourstobill, tbdratesource, tbdrate, tbdamountsource, tbdamount, tbdcode1, tbdcode2, tbdcode3, tbdbillnote, tbdstopwatch, tbdnarrative, tbdid)
                    Select pbbatch, maxrec +1 as Recnbr, 2, 'Y',pbfd, (select prdyear from actngperiod where prdstartdate<=pbfd and prdenddate>=pbfd and prdnbr<>0), 
                    (select prdnbr from actngperiod where prdstartdate<=pbfd and prdenddate>=pbfd and prdnbr<>0),cast('" + Otkpr2.ToString() + @"' as int),cast('" + MatSys.ToString() + @"' as int), 0, matfeesch, null,null, 'Y', 0,'W',0,3, 0, 3, 0, '','','','',0,'', (select max(tbdid) + 1 from timebatchdetail)
                    from  prebillmatter
                    inner join matter on pbmmatter=matsysnbr
                    left outer join tkprrate on tkrfeesch=matfeesch and tkremp=cast('" + Otkpr2.ToString() + @"' as int) 
                    left outer join perstyprate on ptrfeesch=matfeesch and ptrprstyp=(select empprstyp from employee where empsysnbr=cast('" + Otkpr2.ToString() + @"' as int)) , 
(select tbdbatch as pbbatch, max(tbdrecnbr) as maxrec, (select max(pbfdateonbill) from prebillfeeitem ) as pbfd
 from timebatchdetail     where tbdbatch=(select max(pbfutbatch) from prebillfeeitem   )  group by tbdbatch) PBatch 
where  cast('" + Otkpr2.ToString() + @"' as int) is not null and
cast('" + Otkpr2.ToString() + "' as int) > 0 and cast('" + Alloc2.ToString() + @"' as money) <> 0
  and cast('" + Alloc2.ToString() + "' as money) is not null and pbmmatter = cast('" + MatSys.ToString() + @"' as int)
  and  pbmprebill = cast('" + Prebill.ToString() + "' as int)  ";


                    _jurisUtility.ExecuteNonQueryCommand(0, SQL46);

                    string SQL47 = @"insert into timebatchdetail(tbdbatch,tbdrecnbr, tbdrectype, tbdposted, tbddate, tbdprdyear, tbdprdnbr, tbdtkpr, tbdmatter, tbdbudgphase, tbdfeesched, tbdtaskcd, tbdactivitycd,tbdbillableflg,tbdactualhrswrk, tbdhourssource, tbdhourstobill, tbdratesource, tbdrate, tbdamountsource, tbdamount, tbdcode1, tbdcode2, tbdcode3, tbdbillnote, tbdstopwatch, tbdnarrative, tbdid)
                    Select pbbatch, maxrec +1 as Recnbr, 2, 'Y',pbfd, (select prdyear from actngperiod where prdstartdate<=pbfd and prdenddate>=pbfd and prdnbr<>0), 
                    (select prdnbr from actngperiod where prdstartdate<=pbfd and prdenddate>=pbfd and prdnbr<>0),cast('" + Otkpr3.ToString() + @"' as int),cast('" + MatSys.ToString() + @"' as int), 0, matfeesch, null,null, 'Y', 0,'W',0,3, 0, 3, 0, '','','','',0,'', (select max(tbdid) + 1 from timebatchdetail)
                    from  prebillmatter
                    inner join matter on pbmmatter=matsysnbr
                    left outer join tkprrate on tkrfeesch=matfeesch and tkremp=cast('" + Otkpr3.ToString() + @"' as int) 
                    left outer join perstyprate on ptrfeesch=matfeesch and ptrprstyp=(select empprstyp from employee where empsysnbr=cast('" + Otkpr3.ToString() + @"' as int)) , 
(select tbdbatch as pbbatch, max(tbdrecnbr) as maxrec, (select max(pbfdateonbill) from prebillfeeitem  ) as pbfd
 from timebatchdetail     where tbdbatch=(select max(pbfutbatch) from prebillfeeitem   )  group by tbdbatch) PBatch 
where  cast('" + Otkpr3.ToString() + @"' as int) is not null and
cast('" + Otkpr3.ToString() + "' as int) > 0 and cast('" + Alloc3.ToString() + @"' as money) <> 0
  and cast ('" + Alloc3.ToString() + "' as money) is not null and pbmmatter = cast('" + MatSys.ToString() + @"' as int)
  and  pbmprebill = cast('" + Prebill.ToString() + "' as int)  ";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL47);

                    string SQL44 = @"Insert into unbilledtime(utbatch, utrecnbr, utmatter, utbudgphase,utdate, utprdyear, utprdnbr, uttkpr, utfeesched, uttaskcd, utactivitycd, utbillableflg,utactualhrswrk,uthourssource, uthourstobill, utratesource, utrate, utamountsource, utamount, utstdrate, utamtatstdrate, utnarrative,utid, utpostdate, utpostin, utcode1, utcode2, utcode3, utbillnote)
                    select tbdbatch, tbdrecnbr, tbdmatter, tbdbudgphase, tbddate, tbdprdyear, tbdprdnbr, tbdtkpr, tbdfeesched, tbdtaskcd, tbdactivitycd,tbdbillableflg,tbdactualhrswrk, tbdhourssource, tbdhourstobill, tbdratesource, tbdrate,tbdamountsource, tbdamount, tbdrate, tbdamount, tbdnarrative, tbdid, tbddate, -1, tbdcode1, tbdcode2, tbdcode3,tbdbillnote
                    from timebatchdetail,(select pbfutbatch as pbbatch, max(pbfutrecnbr) as maxrec, max(pbfdateonbill) as pbfd
 from prebillfeeitem  where pbfutbatch=(select max(pbfutbatch) from prebillfeeitem)  group by pbfutbatch) PBatch
                    where tbdmatter=cast('" + MatSys.ToString() + @"' as int) and tbdbatch=pbbatch and tbdrecnbr>maxrec";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL44);

                    string SQL121 = @"Insert into prebillfeeitem(pbfprebill, pbfmatter, pbftkpronbill, pbfutbatch, pbfutrecnbr, pbfstatus, pbfdateonbill, pbfhrsonbill, pbfrateonbill, pbfamtonbill, pbfseqonbill)
                Select cast('" + Prebill.ToString() + "' as int), cast('" + MatSys.ToString() + "' as int), uttkpr, utbatch, utrecnbr, 'I',utdate, utactualhrswrk, 0, 0, case when (select max(pbfseqonbill) from prebillfeeitem where pbfprebill=cast('" + Prebill.ToString() + "' as int)) is null then Rank() over (order by utrecnbr) -1 else  (select max(pbfseqonbill)  from prebillfeeitem where pbfprebill=cast('" + Prebill.ToString() + @"' as int)) + Rank() over (order by utrecnbr) end
                from unbilledtime,(select pbfutbatch as pbbatch, max(pbfutrecnbr) as maxrec, max(pbfdateonbill) as pbfd
 from prebillfeeitem    where pbfutbatch=(select max(pbfutbatch) from prebillfeeitem)  group by pbfutbatch) PBatch
                where utmatter=cast('" + MatSys.ToString() + "' as int) and utbatch=pbbatch and utrecnbr>maxrec";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL121);

                    string SQL33 = @"Update prebillfeeitem
                    set pbfamtonbill=case when pbftkpronbill=cast('" + Otkpr1.ToString() + "' as int) then cast('" + Alloc1.ToString() + @"' as money)
                    when pbftkpronbill=cast('" + Otkpr2.ToString() + "' as int) then cast('" + Alloc2.ToString() + @"' as money)
                    when pbftkpronbill=cast('" + Otkpr3.ToString() + "' as int) then cast('" + Alloc3.ToString() + @"' as money) else pbfamtonbill end
                    where pbfprebill=cast('" + Prebill.ToString() + "' as int) and pbfmatter=cast('" + MatSys.ToString() + "' as int) and pbfstatus='I' and pbfhrsonbill=0";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL33);

                    string SQL34 = @"update prebillfeeitem set pbfamtonbill=(pbfamtonbill - cast('" + Rd.ToString() + @"' as money)) from matter where pbfprebill=cast('" + Prebill.ToString() + "' as int) and cast('" + Rd.ToString() + @"' as money)>0 and cast('" + Rd.ToString() + @"' as money)  is not null and pbftkpronbill=cast('" + Otkpr2.ToString() + "' as int) and pbfstatus='I'";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL34);


                    string SQL13 = @"insert into prebillfeerecap(pbfrprebill, pbfrmatter, pbfrtkpronbill, pbfrhrsonbill, pbframtonbill, pbfrppdapplied,pbfrtrustapplied)
                select cast('" + Prebill.ToString() + "' as int),cast('" + MatSys.ToString() + "' as int), cast('" + Otkpr1.ToString() + "' as int),0,0,0,0 from employee where cast('" + Alloc1.ToString() + "' as money)<>0 and cast('" + Alloc1.ToString() + "' as money) is not null and  empsysnbr=cast('" + Otkpr1.ToString() + "' as int) and empsysnbr not in (select pbfrtkpronbill from prebillfeerecap where pbfrprebill=cast('" + Prebill.ToString() + "' as int))";

                    _jurisUtility.ExecuteNonQueryCommand(0, SQL13);


                    string SQL3 = @"insert into prebillfeerecap(pbfrprebill, pbfrmatter, pbfrtkpronbill, pbfrhrsonbill, pbframtonbill, pbfrppdapplied,pbfrtrustapplied)
                select cast('" + Prebill.ToString() + "' as int),cast('" + MatSys.ToString() + "' as int), cast('" + Otkpr2.ToString() + "' as int),0,0,0,0 from employee where cast('" + Alloc2.ToString() + "' as money)<>0 and cast('" + Alloc2.ToString() + "' as money) is not null and  empsysnbr=cast('" + Otkpr2.ToString() + "' as int) and empsysnbr not in (select pbfrtkpronbill from prebillfeerecap where pbfrprebill=cast('" + Prebill.ToString() + "' as int))";

                    _jurisUtility.ExecuteNonQueryCommand(0, SQL3);


                    string SQL4 = @"insert into prebillfeerecap(pbfrprebill, pbfrmatter, pbfrtkpronbill, pbfrhrsonbill, pbframtonbill, pbfrppdapplied,pbfrtrustapplied)
               select cast('" + Prebill.ToString() + "' as int),cast('" + MatSys.ToString() + "' as int), cast('" + Otkpr3.ToString() + "' as int),0,0,0,0 from employee where cast('" + Alloc3.ToString() + "' as money)<>0 and cast('" + Alloc3.ToString() + "' as money) is not null and  empsysnbr=cast('" + Otkpr3.ToString() + "' as int) and empsysnbr not in (select pbfrtkpronbill from prebillfeerecap where pbfrprebill=cast('" + Prebill.ToString() + "' as int))";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL4);



                    string SQL12 = @"update prebillfeerecap set pbframtonbill=amt 
from (select pbfmatter, pbfprebill, pbftkpronbill, sum(pbfamtonbill) as amt from prebillfeeitem inner join matter on pbfmatter=matsysnbr where matbillagreecode='R' group by pbfmatter, pbfprebill, pbftkpronbill)PB
where pbfrmatter=pbfmatter and pbfprebill=pbfrprebill and pbfrtkpronbill=pbftkpronbill and pbfrprebill=cast('" + Prebill.ToString() + "' as int) and pbfrmatter=cast('" + MatSys.ToString() + "' as int)";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL12);







                    string SQL17 = @"update prebillmatter set pbmfeebld=amt 
from (select pbfmatter, pbfprebill,  sum(pbfamtonbill) as amt from prebillfeeitem inner join matter on pbfmatter=matsysnbr where matbillagreecode='R' and pbfprebill=cast('" + Prebill.ToString() + "' as int)  and pbfmatter=cast('" + MatSys.ToString() + @"' as int) group by pbfmatter, pbfprebill)PB where  pbmmatter=pbfmatter and pbmprebill=pbfprebill ";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL17);



                    string SQL20 = "update prebill set pbstatus=4, pbaction=1 where pbsysnbr=cast('" + Prebill.ToString() + "' as int)";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL20);




                }


                string SQLDT = "EXEC sp_MSforeachtable @command1='ALTER TABLE ? CHECK CONSTRAINT ALL'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQLDT);

                string SQLPB = @"select cast(Prebill as varchar(20)) as Prebill, cast(matsysnbr as varchar(10)) as MatterSys,cast(billtobillingatty as varchar(20)) as BTkpr, cast(morig1 as varchar(20)) as OTkpr1,
            cast(case when morig2 is null then '' else morig2 end as varchar(20)) as OTkpr2, 
                cast(case when morig3 is null then '' else morig3 end as  varchar(20)) OTkpr3
            ,convert(varchar(30),cast(isnull(case when (retainertoallocate * OTPct1 * .01)>o1fees then (retainertoallocate * OTPct1 * .01) - o1fees else 0 end,0) as money),1) as OT1Alloc
            ,convert(varchar(30),cast(isnull(case when (retainertoallocate * OTPct2 * .01) >o2fees then (retainertoallocate * OTPct2 * .01) - o2fees else 0 end,0) as money),1) as OT2Alloc
            ,convert(varchar(30),cast(isnull(case when (retainertoallocate * OTPct3 * .01)>o3fees then (retainertoallocate * OTPct3 * .01) - o3fees else 0 end,0) as money),1) as OT3Alloc
           
            from (select prebill as Prebill,billtobillingatty, matsysnbr, matfltfeeorretainer as retainer, sum(totalfees) as TotalFees,  morig1, morig2, morig3, morig4, 
            sum(cast(case when totalfees is null then 0 else totalfees end  as money)) as Total,
            sum(cast(otherfees as money)) as Other,  sum(cast(otkpr1Fees as money))  as O1Fees, ot1, case when cast(otpct1 as int)=0 then '' else  cast(otpct1 as varchar(10)) end  as OTPct1, 
            sum(cast(otkpr2Fees as money)) as O2Fees, ot2, case when cast(otpct2 as int)=0 then '' else  cast(otpct2 as varchar(10)) end  as OTPct2, 
            sum(cast(otkpr3Fees as money)) as O3Fees, ot3, case when cast(otpct3 as int)=0 then '' else  cast(otpct3 as varchar(10)) end   as OTPct3, 
            sum(cast(otkpr4Fees as money)) as O4Fees, ot4, case when cast(otpct4 as int)=0 then '' else  cast(otpct4 as varchar(10)) end  as OTPct4,
            matfltfeeorretainer - sum(cast(case when totalfees is null then 0 else totalfees end as money)) as RetainerToAllocate
            from (select pbmprebill as prebill, utmatter, utamount as TotalFees, case when uttkpr not in (select morigatty from matorigatty where morigmat=matsysnbr) then utamount else null end as OtherFees
            ,case when uttkpr=morig1 then utamount else null end as Otkpr1Fees, case when uttkpr=morig2 then utamount else null end as Otkpr2Fees
            ,case when uttkpr=morig3 then utamount else null end as Otkpr3Fees, case when uttkpr=morig4 then utamount else null end as Otkpr4Fees,morig1, morig2, morig3, morig4, ot1, ot2, ot3, ot4, otpct1, otpct2, otpct3, otpct4, null as Expenses
            from unbilledtime
            inner join prebillmatter on pbmmatter=utmatter
            inner join prebillfeeitem on pbfutbatch=utbatch and pbfutrecnbr=utrecnbr
            inner join matter on pbmmatter=matsysnbr
inner join billto on billtosysnbr = matbillto
            inner join (select morigmat, max(case when rnk=1 then morigatty else '' end) as morig1, max(case when rnk=2 then morigatty else null end) as morig2, max(case when rnk=3 then morigatty else null end) as morig3
            , max(case when rnk=4 then morigatty else null end) as morig4, max(case when rnk=1 then empinitials else null end) as OT1, max(case when rnk=2 then empinitials else null end) as OT2
            , max(case when rnk=3 then empinitials else null end) as OT3, max(case when rnk=4 then empinitials else null end) as OT4, max(case when rnk=1 then cast(morigpcnt as int) else null end) as OTpct1
            , max(case when rnk=2 then cast(morigpcnt as int) else null end) as OTpct2, max(case when rnk=3 then cast(morigpcnt as int) else null end) as OTpct3
            , max(case when rnk=4 then cast(morigpcnt as int) else null end) as OTpct4
            from (select morigmat, morigatty, morigpcnt, empinitials, rank() over (partition by morigmat order by case when billtobillingatty=morigatty then 0 else morigatty end) as rnk
            from matorigatty
            inner join employee on empsysnbr=morigatty
            inner join matter on morigmat=matsysnbr
            inner join billto on matbillto=billtosysnbr)MO
            group by morigmat)MOrig on morigmat=matsysnbr
            where matbillagreecode='R' and matfltfeeorretainer<>0 and matstatusflag='O' and billtobillingatty = " + billToAttyEmpSys
           + @" ) UT
            inner join matter on matsysnbr=utmatter
            inner join client on matclinbr=clisysnbr
            inner join billto on matbillto=billtosysnbr 
            inner join prebill on prebill=pbsysnbr
            where pbstatus<=2 and matbillagreecode='R'  " +
            "  and (cast(billtobillingatty as varchar(20))<>cast(morig1 as varchar(20))  or " +
            " cast(case when morig2 is null then '' else morig2 end as varchar(20))>'') " +
           "  group by matsysnbr, matfltfeeorretainer, ot1, ot2, ot3, ot4, otpct1, otpct2, otpct3, otpct4, prebill, clicode, matcode, matreportingname, clireportingname,morig1, morig2, morig3, morig4,  billtobillingatty " +
            " having matfltfeeorretainer - sum(cast(case when totalfees is null then 0 else totalfees end as money))>0 ) AllocTbl order by prebill";

                DataSet myRS = _jurisUtility.RecordsetFromSQL(SQLPB);

                dataGridView1.AutoGenerateColumns = true;

                dataGridView1.DataSource = myRS.Tables[0];


                toolStripStatusLabel.Text = "Processing Completed.";
                statusStrip.Refresh();
            }
            else
            {
                string SQLDT = "EXEC sp_MSforeachtable @command1='ALTER TABLE ? CHECK CONSTRAINT ALL'";
                _jurisUtility.ExecuteNonQueryCommand(0, SQLDT);


                toolStripStatusLabel.Text = "No Records to Process.";
                statusStrip.Refresh();
            }

            Cursor.Current = Cursors.Default;
            Application.DoEvents();
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
            if (File.Exists(filePathName))
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

        private void btn_Prebill_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();
            toolStripStatusLabel.Text = "Selecting Prebills....";
            statusStrip.Refresh();
            string SQLPB = @"select Prebill, ClientNumber, ClientName, MatterName, Retainer,convert(varchar(30),TotalFees,1) as Total,
convert(varchar(30),Other,1) as OtherTkprs, ot1 as OTkpr1, otpct1 as OTPct1,   convert(varchar(30),cast(case when o1fees=0 then null else o1fees end as money),1)as OTFees1,
 ot2 as OTkpr2, otpct2 as OTPct2,convert(varchar(30),cast(case when o2fees=0 then null else o2fees end as money),1) as OTFees2 , 
 ot3 as OTkpr3, Otpct3 as OTPct3 ,convert(varchar(30),cast(case when o3fees=0 then null else o3Fees end as money),1) as OTFees3, 
 convert(varchar(30),retainertoallocate,1) as RetainerBalance
,convert(varchar(30),cast(case when retainertoallocate>0 then (retainertoallocate * OTPct1 * .01)  else 0 end as money),1) as OTAlloc1
,convert(varchar(30),cast(case when retainertoallocate>0 then (retainertoallocate * OTPct2 * .01)  else 0 end as money),1) as OTAlloc2
,convert(varchar(30),cast(case when retainertoallocate>0 then (retainertoallocate * OTPct3 * .01)  else 0 end as money),1) as OTAlloc3
from (select prebill as Prebill, CliCode  as ClientNumber, dbo.jfn_formatmattercode(matcode) as MatterNumber, clireportingname as ClientName, 
matreportingname as MatterName, matfltfeeorretainer as retainer, sum(totalfees) as TotalFees,
sum(cast(case when totalfees is null then 0 else totalfees end as money)) as Total,
sum(cast(otherfees as money)) as Other, sum(cast(otkpr1Fees as money))  as O1Fees, ot1, case when cast(otpct1 as int)=0 then '' else  cast(otpct1 as varchar(10)) end  as OTPct1, 
sum(cast(otkpr2Fees as money)) as O2Fees, ot2, case when cast(otpct2 as int)=0 then '' else  cast(otpct2 as varchar(10)) end  as OTPct2, 
sum(cast(otkpr3Fees as money)) as O3Fees, ot3, case when cast(otpct3 as int)=0 then '' else  cast(otpct3 as varchar(10)) end   as OTPct3, 
sum(cast(otkpr4Fees as money)) as O4Fees, ot4, case when cast(otpct4 as int)=0 then '' else  cast(otpct4 as varchar(10)) end  as OTPct4,
matfltfeeorretainer - sum(cast(case when totalfees is null then 0 else totalfees end  as money)) as RetainerToAllocate
from (select pbmprebill as prebill, utmatter, utamount as TotalFees, 
case when uttkpr not in (select morigatty from matorigatty where morigmat=matsysnbr) then utamount else null end as OtherFees
, case when uttkpr=morig1 then utamount else null end as Otkpr1Fees
, case when uttkpr=morig2 then utamount else null end as Otkpr2Fees
, case when uttkpr=morig3 then utamount else null end as Otkpr3Fees
, case when uttkpr=morig4 then utamount else null end as Otkpr4Fees, ot1, ot2, ot3, ot4, otpct1, otpct2, otpct3, otpct4
from unbilledtime
inner join prebillmatter on pbmmatter=utmatter
inner join matter on pbmmatter=matsysnbr
 inner join prebillfeeitem on pbfutbatch=utbatch and pbfutrecnbr=utrecnbr
inner join (select morigmat, max(case when rnk=1 then morigatty else '' end) as morig1
, max(case when rnk=2 then morigatty else null end) as morig2
, max(case when rnk=3 then morigatty else null end) as morig3
, max(case when rnk=4 then morigatty else null end) as morig4
, max(case when rnk=1 then empinitials else null end) as OT1
, max(case when rnk=2 then empinitials else null end) as OT2
, max(case when rnk=3 then empinitials else null end) as OT3
, max(case when rnk=4 then empinitials else null end) as OT4
, max(case when rnk=1 then cast(morigpcnt as int) else null end) as OTpct1
, max(case when rnk=2 then cast(morigpcnt as int) else null end) as OTpct2
, max(case when rnk=3 then cast(morigpcnt as int) else null end) as OTpct3
, max(case when rnk=4 then cast(morigpcnt as int) else null end) as OTpct4
from (select morigmat, morigatty, morigpcnt, empinitials, rank() over (partition by morigmat order by morigatty) as rnk
from matorigatty
inner join employee on empsysnbr=morigatty)MO
group by morigmat)MOrig on morigmat=matsysnbr
) UT
inner join matter on matsysnbr=utmatter
inner join client on matclinbr=clisysnbr
inner join prebill on prebill=pbsysnbr
where pbstatus<=2 and pbbillto in (select billtosysnbr from billto where billtobillingatty = " + billToAttyEmpSys + ") " +
" group by prebill , CliCode  , dbo.jfn_formatmattercode(matcode) , clireportingname , matreportingname , matfltfeeorretainer, ot1, ot2, ot3, ot4, otpct1, otpct2, otpct3, otpct4 " + 
" ) AllocTbl order by prebill";

            DataSet myRS = _jurisUtility.RecordsetFromSQL(SQLPB);
         

                dataGridView1.AutoGenerateColumns = true;

                dataGridView1.DataSource = myRS.Tables[0];

                toolStripStatusLabel.Text = "Ready to Process Prebills....";
                statusStrip.Refresh();

                Cursor.Current = Cursors.Default;
                Application.DoEvents();
            }

        


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            billToAttyEmpSys = comboBox1.SelectedValue.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string s2 = "update prebill set pbstatus=2, pbaction=0 where pbsysnbr=" + tbPrebill.Text.ToString();

            _jurisUtility.ExecuteNonQuery(0, s2);

            MessageBox.Show(tbPrebill.ToString() + " has been reset to ready to edit status.","Prebill Status", MessageBoxButtons.OK);

            string sql = "SELECT distinct empname, BillToBillingAtty FROM PreBill " +
                      "inner join billto on billtosysnbr = pbbillto " +
                      "inner join employee on empsysnbr = BillToBillingAtty " +
                      " inner join prebillmatter on pbmprebill = pbsysnbr " +
                      " inner join matter on matsysnbr = pbmmatter " +
                      " where pbstatus <= 2 and matbillagreecode = 'R' and matfltfeeorretainer<>0 and matstatusflag='O' ";
            DataSet emp = _jurisUtility.RecordsetFromSQL(sql);
            if (emp == null || emp.Tables[0].Rows.Count == 0)
            {
                MessageBox.Show("There are no prebills to process", "No processing", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }



            else
            {
                comboBox1.ValueMember = "BillToBillingAtty";
                comboBox1.DisplayMember = "empname";
                comboBox1.DataSource = emp.Tables[0];
            }

        }
    
    }
}

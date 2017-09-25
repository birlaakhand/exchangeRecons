using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using System.Data;
using Recon.Tools.HTML;
using System.Reflection;
using mshtml;
using CefSharp;
using System.Diagnostics;
using CefSharp.MinimalExample.Wpf;
using System.Timers;
using System.Windows.Threading;
//using System.Web.Ui>

namespace ExchangeRecon
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dictionary<string, ReconQueue> ICEqueues, CMEqueues;
        DispatcherTimer ICEScheduler, ICEReconTimer, ICEOLImporterTimer;
        DBConnector db;
        string ICEOLQuery;
        public MainWindow()
        {
            InitializeComponent();
            //testQmagic();
            initializeExchanges();
            //tempGrid2();
        }

        public void initializeExchanges()
        {
            //Initialize ICE
            ReconQueue ICEqueuesData = new ReconQueue(
                @"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\Queues.csv");
            webPagePreview.DownloadHandler = new DownloadHandler();
            webPagePreview.Address = @"https://www.theice.com/reports/DealReport.shtml?";

            ICEqueues = new Dictionary<string, ReconQueue>();
            foreach (DataRow r in ICEqueuesData.QData.Rows)
            {
                ICEqueues.Add(r["QueueName"].ToString(), new ReconQueue(r["AttachedFile"].ToString()));
                Console.WriteLine(r["QueueName"].ToString() + " Queue added");
            }
            QueueSelector.ItemsSource = ICEqueues;
            QueueSelector.DisplayMemberPath = "Key";
            QueueSelector.SelectedValuePath = "Key";
            QueueSelector.SelectedItem = ICEqueues["Queues"];

            ICEOLQuery = "SELECT (NVL(SUBSTR(a.reference , 0, INSTR(a.reference , '.')-1), a.reference ))  AS ICEdealID," +
                        " a.tran_num, a.deal_tracking_num, a.reference, a.trade_date, a.position, price, '' as Contract " +
                        " FROM ceiolfprod.ab_tran a   inner join ceiolfprod.header h on a.ins_num = h.ins_num " +
                        " WHERE book IN ( 'ICEGateway','CommStpGateway 1:ICE-OTC' ) AND trade_date=  '%TRADEDATE%' " +
                        " AND trade_flag = 1 AND tran_type NOT IN (27, 41) AND tran_status IN (2, 3) AND asset_type = 2";
            
            ICEScheduler = new DispatcherTimer();
            ICEScheduler.Tick += new EventHandler(ICEScheduleMethod);
            ICEScheduler.Interval = new TimeSpan(0, 5, 0);

            ICEReconTimer = new DispatcherTimer();
            ICEReconTimer.Tick += new EventHandler(reconSchedulerMethod);
            ICEReconTimer.Interval = new TimeSpan(0, 0, 6);

            ICEOLImporterTimer = new DispatcherTimer();
            ICEOLImporterTimer.Tick += new EventHandler(ICEOLImportSchedulerMethod);
            ICEOLImporterTimer.Interval = new TimeSpan(0, 0, 6);

            ICEReconStatus.Content = "Not Started";
            ICEOLImporterStatus.Content = "Not Started";
            ICEDownloaderStatus.Content = "Not Started";
            ICESchedulerStatus.Content = "Not Started";

            //Initialize CME
            ReconQueue CMEqueuesData = new ReconQueue(
            @"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\Queues.csv");
            CMEwebPagePreview.DownloadHandler = new DownloadHandler();
            CMEwebPagePreview.Address = @"https://login.cmegroup.com/sso/ssologin.action";

            CMEqueues = new Dictionary<string, ReconQueue>();
            foreach (DataRow r in ICEqueuesData.QData.Rows)
            {
                CMEqueues.Add(r["QueueName"].ToString(), new ReconQueue(r["AttachedFile"].ToString()));
                Console.WriteLine(r["QueueName"].ToString() + " Queue added");
            }
            CMEQueueSelector.ItemsSource = CMEqueues;
            CMEQueueSelector.DisplayMemberPath = "Key";
            CMEQueueSelector.SelectedValuePath = "Key";
            CMEQueueSelector.SelectedItem = CMEqueues["Queues"];


        }

        public DataTable OracleTableSanitizer(DataTable oracle)
        {
            foreach (DataRow dr in oracle.Rows)
            {
                dr["ICEDEALID"] = System.Text.RegularExpressions.Regex.Replace(dr["ICEDEALID"].ToString(), "[A-Za-z ]", "");
            }
            return oracle;
        }

        public void testQmagic()
        {
            var a = new HTMLParser(@"C:\Users\Akhand\Downloads\DealReport.xls").Process().ToList();
            Console.WriteLine("Num of tables extracted : " + a.Count);

            var data = new DataTable("ICE Data");
            MergeTables(a, data);

            var res = from t1 in data.AsEnumerable()
                      where t1.Field<String>("TableName") == "Cleared Deals"
                      select t1;
            var res2 = data.Select(@"TableName like '%Cleared Deals%' and RowNum in (2,4,6)");
            var res3 = data.Select(@"TableName like '%Cleared Deals%' and RowNum in (2,4,7)");
            var res4 = from q1 in res2
                       join q2 in res3 on q1["RowNum"] equals q2["RowNum"]
                       where q1.Field<String>("Deal ID").Contains("24") && q1.Field<String>("RowNum")=="4"
                       select q1;
            var res5 = res2.Except(res3);
            //dataGridView1.DataContext = res2.CopyToDataTable();
            string testFile = @"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\Queues.csv";
            ReconQueue queuesData = new ReconQueue(
                testFile);
            Console.WriteLine("Idhar Hu");
            DataView dgv = new DataView(queuesData.QData);
            //listView.ItemsSource = queuesData.QData.Rows;
            //listView.ItemsSource = dgv;
            //listView.it
            dataGridView1.ItemsSource = dgv;
            Console.WriteLine("Ab Yahan");
            Dictionary<string, ReconQueue> queues = new Dictionary<string, ReconQueue>();
            Console.WriteLine("Le yahan bhi aa gaya");
            foreach (DataRow r in queuesData.QData.Rows)
            {
                queues.Add(r["QueueName"].ToString(), new ReconQueue(r["AttachedFile"].ToString()));
                Console.WriteLine(r["QueueName"].ToString() + " Queue added");
            }

            ReconQueue comparisons = new ReconQueue(
                @"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\Comparer3.csv");
            foreach(DataRow comp in comparisons.QData.Rows)
                queryTrials(comp, queues);
            Console.WriteLine("Analysis done");
            foreach (KeyValuePair<string,ReconQueue> q in queues)
            {
                q.Value.ToCSV();
            }
            Console.WriteLine("Sab ho gaya");
            //queryTrials();

        }

        private void queryTrials(DataRow inp, Dictionary<string,ReconQueue> rData)
        {
            Console.WriteLine(inp["Queue1"].ToString());
            Console.WriteLine(inp["Queue2"].ToString());
            DataTable q1 = rData[inp["Queue1"].ToString()].QData;
            DataTable q2 = rData[inp["Queue2"].ToString()].QData;
            DataTable target = rData[inp["TargetQueue"].ToString()].QData;
            //string p1 = q1.PrimaryKey[0].ColumnName, p2 = q2.PrimaryKey[0].ColumnName, pT = target.PrimaryKey[0].ColumnName;
            string a1 = inp["Queue1 Action"].ToString(), a2 = inp["Queue2 Action"].ToString(), aT = inp["Target Action"].ToString();

            string match = inp["Join Type"].ToString();
            string StructMatch = inp["Queue Structure"].ToString();
            string c1 = inp["Params"].ToString().Split(';')[0], c2 = inp["Params"].ToString().Split(';')[1];

            DataColumn[] dc1, dc2;
            string s = inp["Select"].ToString();
            string[] s1 = s.Split(';')[0].Split(','), s2 = s.Split(';')[1].Split(',');
            int l1 = s1.Count(), l2 = s2.Count();

            if (l1 == 1 && s1[0].Equals("*"))
            {
                dc1 = new DataColumn[q1.Columns.Count];
                q1.Columns.CopyTo(dc1, 0);
            }
            else if (l1 == 1 && s1[0].Equals("NA"))
            {
                dc1 = new DataColumn[0];
            }
            else
            {
                dc1 = new DataColumn[l1];
                for (int i = 0; i < l1; i++)
                    dc1[i] = q1.Columns[s1[i]];
            }
            if (l2 == 1 && s2[0].Equals("*"))
            {
                dc2 = new DataColumn[q2.Columns.Count];
                q2.Columns.CopyTo(dc2, 0);
            }
            else if (l2 == 1 && s2[0].Equals("NA"))
            {
                dc2 = new DataColumn[0];
            }
            else
            {
                dc2 = new DataColumn[l2];
                for (int i = 0; i < l2; i++)
                    dc2[i] = q2.Columns[s2[i]];
            }


            var t = new DataTable("Temp Results");
            switch (StructMatch)
            {
                case "Same" :                   //Table Structures are same
                    switch (match)
                    {
                        case "Outer Exclude":         //Records in Q1, but not in Q2
                            t = q1.Clone();
                            if (q1.Rows.Count > 0 && q2.Rows.Count > 0)
                            {
                                var rows = q1.AsEnumerable().Except(q2.AsEnumerable(), DataRowComparer.Default);
                                if (rows.Count() > 0)
                                    t = rows.CopyToDataTable();
                            }
                            else
                                t = q1;
                            break;
                    }
                    break;
                case "Diff":
                    switch (match)
                    {
                        case "Outer Exclude":
                            t = q1.AsEnumerable().Except(q2.AsEnumerable()).CopyToDataTable();
                            break;
                        case "Inner":
                            t = compareTablesCustom(q1, q2, q1.Columns[c1], q2.Columns[c2],dc1,dc2);
                            break;
                        case "InnerOld":
                            var p = from d1 in q1.AsEnumerable()
                                    from d2 in q2.AsEnumerable()
                                    where d1.Field<string>("Ref ID").Contains(d2.Field<string>("Deal ID"))
                                    //select d1;
                                    select new { d1, d2 };
                            t = p.Cast<DataRow>().CopyToDataTable();
                            break;
                    }
                    break;
            }

            switch(a1 + "_" + a2)
            {
                case "Remove_None":
                    q1 = modifyParentTables(q1,t, q1.Columns[c1],dc1);
                    break;
                case "None_Remove":
                    q2 = modifyParentTables(q2,t, q2.Columns[c2], dc2);
                    break;
                case "Remove_Remove":
                    q1 = modifyParentTables(q1,t, q1.Columns[c1], dc1);
                    q2 = modifyParentTables(q2,t, q2.Columns[c2], dc2);
                    break;
            }

            switch (aT)
            {
                case "Overwrite":
                    target = t;
                    break;
                case "Append": 
                    foreach (DataRow r in t.Rows)
                        target.Rows.Add(r.ItemArray);
                    break;
            }

            rData[inp["Queue1"].ToString()].QData = q1;
            rData[inp["Queue2"].ToString()].QData = q2;
            rData[inp["TargetQueue"].ToString()].QData = target;
            Console.Write("\n");
            foreach (KeyValuePair<string, ReconQueue> q in rData)
            {
                Console.Write(q.Key.ToString() + "["+ q.Value.QData.Rows.Count + "] ; ");
            }
            Console.Write("\n");
            GC.Collect();
        }

        public DataTable modifyParentTables(DataTable P, DataTable T, DataColumn c, DataColumn[] s)
        {
            DataTable res = P.Clone();
            res.Clear();
            Console.WriteLine("Size of P = " + P.Rows.Count + ", T = " + T.Rows.Count + "; Num Data Cols = " + s.Count());
            foreach (DataRow r in P.Rows)
            {
                bool rowMatch = false;
                foreach (DataRow tr in T.Rows)
                {
                    foreach(DataColumn dc in s)
                    {
                        //Console.WriteLine(r[dc.ColumnName].ToString());
                        //Console.WriteLine(tr[dc.ColumnName].ToString());
                        if (r[dc.ColumnName] != tr[dc.ColumnName])
                        {
                            rowMatch = false;
                            break;
                        }
                        else
                            rowMatch = true;
                    }
                    if (rowMatch == true)
                        break;
                }
                if (rowMatch == false)
                    res.Rows.Add(r.ItemArray);
            }
            return res;
        }

        public DataTable compareTablesCustom(DataTable d1, DataTable d2, DataColumn c1, DataColumn c2, DataColumn[] s1, DataColumn[] s2)
        {
            DataTable res = new DataTable("Result");
            foreach (DataColumn c in s1)
                res.Columns.Add(c.ColumnName);
            res.Clear();
            foreach (DataColumn c in s2)
                res.Columns.Add(c.ColumnName);
            res.Clear();

            
            foreach (DataRow r1 in d1.Rows)
            {
                foreach(DataRow r2 in d2.Rows)
                {
                    if (r1.Field<string>(c1.ColumnName).Length>0 && r2.Field<string>(c2.ColumnName).Contains(r1.Field<string>(c1.ColumnName)))
                    {
                        DataRow r = res.Rows.Add();
                        foreach (DataColumn c in s1)
                            r[c.ColumnName] =r1[c.ColumnName];
                        foreach (DataColumn c in s2)
                            r[c.ColumnName] = r2[c.ColumnName];
                    }
                        
                }
            }
            return res;
        }

        public void tempGrid2()
        {
            /*
            Load Exchange Pages
            Download the results at certain intervals
            */

            var a = new HTMLParser(@"C:\Users\Akhand\Downloads\DealReport.xls").Process().ToList();
            Console.WriteLine("Num of tables extracted : " + a.Count);

            var data = new DataTable("ICE Data");
            MergeTables(a, data);
            
            var res = from t1 in data.AsEnumerable()
                      where t1.Field<String>("TableName") == "Cleared Deals"
                      select t1;
            var res2 = data.Select(@"TableName = 'Cleared Deals'");
            webPagePreview.Address = @"https://www.theice.com/reports/DealReport.shtml?";
            
            //DataView dgv = new DataView(res2.CopyToDataTable());
            //dataGridView1.ItemsSource = dgv;
            //dataGridView1.DataContext = res2.CopyToDataTable();

            ReconQueue t = new ReconQueue(data, @"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\CollatedData.csv");
            ReconQueue queuesData = new ReconQueue(
                @"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\Queues.csv");
            //dataGridView1.ItemsSource = new DataView(queuesData.QData);
            webPagePreview.DownloadHandler = new DownloadHandler();
            
            //bc.OnBeforeDownload((IBrowser)webPagePreview,DownloadItem)
            //queuesData.QData.RowChanged += ReconQueue.Equals();

            ICEqueues = new Dictionary<string, ReconQueue>();
            foreach (DataRow r in queuesData.QData.Rows)
            {
                ICEqueues.Add(r["QueueName"].ToString(), new ReconQueue(r["AttachedFile"].ToString()));
                Console.WriteLine(r["QueueName"].ToString() + " Queue added");
            }
            QueueSelector.ItemsSource = ICEqueues;
            QueueSelector.DisplayMemberPath = "Key";
            QueueSelector.SelectedValuePath = "Key";
            QueueSelector.SelectedItem = ICEqueues["Queues"];
            //webPagePreview.Address = "www.theice.com";
            Debug.WriteLine(webPagePreview.Address);
            //webPagePreview.DownloadHandler = IDownloadHandler p;

            
            /*
            1. GET Data : ICE, CME, OL
            2. Store this data to Queue Objects
            3. Update Queues File
            4. Read and compare Queues
            */
            
            /*
            ReconQueue t = new ReconQueue();
            Console.WriteLine("********************REACHED HERE***********************");
            //t.filepath = @"C:\Users\Akhand\Downloads\DealReport.xlsx";
            //t.sheetname = "Collated";
            t.QData.Rows[2]["Type"] = "Modified";
            Console.WriteLine("********************LAST LAP***********************");
            t.adp.Update(t.dsXLS);
            //t.adp.Update(t.QData);
            //t.adp.Update()
            Console.WriteLine("********************PHEWWW***********************");
            */
        }

        public void MergeTables(List<DataTable> TableList, DataTable CollatedData)
        { 
            // Generate a combined table with added columns of Table Name, Row numbers
            CollatedData.Columns.Add("TableName", typeof(System.String));
            //CollatedData.Columns.Add("PK", typeof(System.String));
            CollatedData.Columns.Add("RowNum", typeof(System.Int16));

            int i = 0;
            foreach (DataTable table in TableList )
            {
                string tName = table.TableName;
                foreach (DataRow row in table.Rows)
                {
                    CollatedData.Rows.Add();
                    CollatedData.Rows[i]["TableName"] = tName;
                    CollatedData.Rows[i]["RowNum"] = i+1;
                    
                    foreach (DataColumn DataCell in table.Columns)
                    {
                        //Adding Column in the Collated Table if it doesn't already exist
                        if (!CollatedData.Columns.Contains(DataCell.ColumnName))
                        {
                            CollatedData.Columns.Add(DataCell.ColumnName,DataCell.DataType);
                        }
                        CollatedData.Rows[i][DataCell.ColumnName] = row[DataCell.ColumnName];
                    }
                    //CollatedData.Rows[i]["PK"] = CollatedData.Rows[i]["Deal ID"] + "_" + CollatedData.Rows[i]["Leg ID"] + "_" + CollatedData.Rows[i]["Orig ID"] + "_" + CollatedData.Rows[i]["Link ID"];
                    i++;
                }
            }
            //CollatedData.PrimaryKey = new DataColumn[] { CollatedData.Columns["PK"] };
        }

        public void AddPKtoICE(DataTable data)
        {
            data.Columns.Add("PK", typeof(System.String));
            foreach(DataRow r in data.Rows)
            {
                r["PK"] = r["Deal ID"] + "_" + r["Leg ID"] + "_" + r["Orig ID"] + "_" + r["Link ID"];
            }
            data.PrimaryKey = new DataColumn[] { data.Columns["PK"] };
        }

        public void tempGrid()
        {
            string XLS_FILE_NAME_AND_PATH_HERE = "C:\\Users\\Akhand\\Downloads\\DealReport.xlsx",
                SHEETNAME_HERE = "DealReport";
            OleDbConnection con = new OleDbConnection(
                "provider=Microsoft.Jet.OLEDB.4.0;data source=" + XLS_FILE_NAME_AND_PATH_HERE
                + ";Extended Properties=Excel 8.0;");

            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [" + SHEETNAME_HERE + "$A1:AD5378] where [F2] <> null");
            OleDbDataAdapter adp = new OleDbDataAdapter(stbQuery.ToString(), con);

            DataSet dsXLS = new DataSet();
            adp.Fill(dsXLS);

            int row;
            foreach (DataTable t in dsXLS.Tables)
            {
                t.Columns.Add("RowNum", typeof(Int32));
                row = 1;
                foreach (DataRow r in t.Rows)
                {
                    r["RowNum"] = row++;
                }
            }

            DataView dvEmp = new DataView(dsXLS.Tables[0]);

            dataGridView1.ItemsSource = dvEmp;
        }

        private void Window_Loaded(object sender, NavigationEventArgs e)
        {
            /*
            //Console.WriteLine(webPagePreview.Document.ToString());
            Console.WriteLine("abdn" + sender.ToString());
            dynamic doc = webPagePreview.Document;
            if (doc != null)
            {
                Console.WriteLine("1234");
                string htmlText = doc.documentElement.InnerHtml;
                Console.WriteLine(htmlText.Length + "==================" + htmlText.Substring(0,20) );
                var x = (IHTMLDocument3)webPagePreview.Document;
                IHTMLDocument3 abc = (IHTMLDocument3)webPagePreview.Document;
                
                Console.WriteLine("...........");// + htmlText.Substring(htmlText.IndexOf("<a href=\"#login\""), 50));
                IHTMLElementCollection y = x.getElementsByTagName("A");
                
                foreach(IHTMLAnchorElement a in y)
                {
                    string h = a.href;
                    if (h.IndexOf("#login") > 0)
                    {
                        Console.WriteLine(a.href + "_________" + a.ToString());
                        IHTMLElement t = (IHTMLElement) a;
                        //t.click();
                        
                    }
                        
                }
            }
            */
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            /*
            When Start Button is pressed, Evaluate if scheduler is ON.
            If ON, put method A on repeat
            Else execute Method A once

            METHOD A
            Check if ICE Downloader is enabled
            If so, trigger the Javascript
            Else, do nothing

            Check if OL-Importer is enabled
            If so, trigger the DB Query
            Else, do nothig

            Check if Recon is enabled
            If so, run recon with a delay
            If not, do nothing

            */
            if (ICEStartButton.Content.ToString() == "Stop")
            {
                ICEStartButton.Content = "Start";
                SchedulerSwitch.IsEnabled = true;
                ReconSwitch.IsEnabled = true;
                olImporterSwitch.IsEnabled = true;
                ICEDownloaderSwitch.IsEnabled = true;
                ICEScheduler.Stop();
            }
            else
            {
                if (ICEStartButton.Content.ToString() == "Start")
                {
                    ICEStartButton.Content = "Stop";
                    SchedulerSwitch.IsEnabled = false;
                    ReconSwitch.IsEnabled = false;
                    olImporterSwitch.IsEnabled = false;
                    ICEDownloaderSwitch.IsEnabled = false;
                }
                if (SchedulerSwitch.Value == 1)
                {
                    ICEScheduleSteps();         //Execute for 1st time now itself
                    ICEScheduler.Start();       //Scheduled job would execute in a while
                }
                else
                {
                    ICEScheduleSteps();         //Scheduled task run once without schedule
                }
            }
                
        }

        private void ICEScheduleMethod(object sender, EventArgs e)
        {
            ICEScheduleSteps();
        }

        private void ICEScheduleSteps()
        {
            ICESchedulerStatus.Content = "Running";
            if (ICEDownloaderSwitch.Value == 1)
            {
                downLoadStarter();
                ICEReconTimer.Interval = new TimeSpan(0, 1, 0);
            }
            if (olImporterSwitch.Value == 1)
            {
                ICEOLImporterTimer.Start();
                ICEReconTimer.Interval = new TimeSpan(0, 1, 0);
            }
            //Running it irrespective of switch coz we need to merge the ICE Files and update the queues. Switch check is in the scheduler method
            ICEReconTimer.Start();
            DateTime now = TimeZoneInfo.ConvertTimeFromUtc(System.DateTime.UtcNow, TimeZoneInfo.Local);
            ICESchedulerStatus.Content = "Last executed at " + now.ToString("dd-mmm-yy hh:mm:ss");
        }

        private void reconSchedulerMethod(object sender, EventArgs e)
        {
            if (ICEDownloaderSwitch.Value == 1)
            {
                var a = new HTMLParser(@"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\Download Data\ICE_CBNA.xls").Process().ToList();
                var b = new HTMLParser(@"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\Download Data\ICE_CGML.xls").Process().ToList();
                foreach (DataTable t in b)
                    a.Add(t);
                Console.WriteLine("Num of tables extracted : " + a.Count);

                var data = new DataTable("ICE Data");
                MergeTables(a, data);
                AddPKtoICE(data);
                ICEqueues["ICE Raw Collated"].QData = data;
                ICEqueues["ICE Raw Collated"].ToCSV();
            }
            Console.WriteLine("Chal gaya hu");
            ICEReconTimer.Stop();
            if (ReconSwitch.Value == 1)
            {
                ICEReconStatus.Content = "Running";
                // Add the Recon Method
                ReconQueue comparisons = ICEqueues["Comparisons"];
                foreach (DataRow comp in comparisons.QData.Rows)
                    queryTrials(comp, ICEqueues);
                Console.WriteLine("Analysis done");
                foreach (KeyValuePair<string, ReconQueue> q in ICEqueues)
                {
                    q.Value.ToCSV();
                }
                DateTime now = TimeZoneInfo.ConvertTimeFromUtc(System.DateTime.UtcNow, TimeZoneInfo.Local);
                ICEReconStatus.Content = "Last executed at " + now.ToString("dd-mmm-yy hh:mm:ss");
            }
        }

        private void ICEOLImportSchedulerMethod(object sender, EventArgs e)
        {
            ICEOLImporterTimer.Stop();
            ICEOLImporterStatus.Content = "Running";
            using (db = new DBConnector())
            {
                DateTime now = TimeZoneInfo.ConvertTimeFromUtc(System.DateTime.UtcNow, TimeZoneInfo.Local);
                string date = now.ToString("dd-MMM-yyyy");
                ICEqueues["OL Raw"].QData = OracleTableSanitizer(db.MakeQuery(ICEOLQuery.Replace("%TRADEDATE%", date), "OL Raw")); ;
                ICEqueues["OL Raw"].ToCSV();
                ICEOLImporterStatus.Content = "Last Updated at " + now.ToString("dd-mmm-yy hh:mm:ss");
            }
            if (ICEOLImporterStatus.Content.ToString() == "Running")
                ICEOLImporterStatus.Content = "Failed";
        }

        private void downLoadStarter()
        {
            //Console.Out;
            //string src =  IBrowser.GetSourceAsync(webPagePreview);
            //webPagePreview.ViewSource();
            //webPagePreview.GetMainFrame().ViewSource();
            if(ReconSwitch.Value==1) ICEReconTimer.Start();
            ICEDownloaderStatus.Content = "Running";
            string cmd = @"document.getElementById('companyId').selectedIndex=0; ";
            cmd += @"var a = document.getElementsByTagName('a'); ";
            cmd += @"var c; for(i=0;i<a.length;i++) if(a[i].href=='https://www.theice.com/reports/DealReport.shtml?excel=') c=a[i]; c.click(); ";
            //cmd += @"alert(a[3].href);"; 
            //cmd += @"document.getElementById('companyId').selectedIndex=1; c.click()";
            //cmd += @"var a = document.getElementsByTagName('a'); ";
            //cmd += "for(i=0;i<a.length;i++) if(a[i].href==\"https://www.theice.com/reports/DealReport.shtml?excel=\") a[i].click(); ";

            cmd = @"var a = document.getElementsByTagName('a'); var c; for (i = 0; i < a.length; i++) if (a[i].href == 'https://www.theice.com/reports/DealReport.shtml?excel=') c = a[i];";
            cmd += @"document.getElementById('companyId').selectedIndex = 0; c.click(); setTimeout(function(){ document.getElementById('companyId').selectedIndex = 1; c.click(); },15000);";
            IFrame frame = webPagePreview.GetMainFrame();
            frame.ExecuteJavaScriptAsync(cmd);
            frame.ExecuteJavaScriptAsync("document.getElementById('companyId').selectedIndex = 1;");
            //frame.ExecuteJavaScriptAsync("alert('ExecuteJavaScript works!');");
            //frame->GetURL(), 0);
            /*
            var a = document.getElementsByTagName('a'); var c; for(i=0;i<a.length;i++) if(a[i].href=='https://www.theice.com/reports/DealReport.shtml?excel=') c=a[i];  setInterval(function(){
            document.getElementById('companyId').selectedIndex=0; c.click(); console.log("0 Done"); setTimeout(function(){document.getElementById('companyId').selectedIndex=1; c.click();},5000); 
            }, 600000);
            */

            Debug.WriteLine("..........." + webPagePreview.Content.ToString());
            DateTime now = TimeZoneInfo.ConvertTimeFromUtc(System.DateTime.UtcNow, TimeZoneInfo.Local);
            ICEDownloaderStatus.Content = "Last executed at " + now.ToString("dd-mmm-yy hh:mm:ss");
            //Console.WriteLine("..........." + webPagePreview.Content.ToString());
        }
        
        private void slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (sender.GetType() == typeof(Slider))
            {
                Slider s = (Slider)sender;
                var bc = new BrushConverter();
                if(s.Value==0)
                    s.Background = (Brush)bc.ConvertFrom("#FFD62222");
                else
                    s.Background = (Brush)bc.ConvertFrom("#FF54D622");
                if (sender.GetHashCode() == SchedulerSwitch.GetHashCode())
                {
                    if (s.Value == 0)
                        ICEStartButton.Content = "Run";
                    else
                        ICEStartButton.Content = "Start";
                }
                s.Background.Opacity = 0.5;
            }
        }

        private void QueueSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            dataGridView1.ItemsSource = new DataView(ICEqueues[QueueSelector.SelectedValue.ToString()].QData);
        }

        /*
private void StartButton_Click(object sender, RoutedEventArgs e)
{
var x = (IHTMLDocument3)webPagePreview.Document;
IHTMLElementCollection srcPages = x.getElementsByName("_sourcePage");
IHTMLElementCollection FP = x.getElementsByName("__fp");
foreach (HTMLInputElement srcPage in srcPages)
{
Console.WriteLine(srcPage.getAttribute("value").ToString());
}
foreach (HTMLInputElement f in FP)
{
Console.WriteLine(f.getAttribute("value").ToString());
}
}
*/
    }
}

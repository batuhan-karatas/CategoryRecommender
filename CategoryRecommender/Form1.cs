using CategoryRecommender.dbModels;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Linq;
using System.Threading;
using System.Data;
using System.Data.SqlClient;

using Excel = Microsoft.Office.Interop.Excel;

namespace CategoryRecommender
{
    public partial class TubitakUzayApp : Form
    {
        public TubitakUzayApp()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            MySql.Data.MySqlClient.MySqlConnection conn;
            string myConnectionString;
            myConnectionString = "server=localhost;uid=root;" + "pwd=1234;database=testdb";
            try
            {
                // connetting MySql
                conn = new MySql.Data.MySqlClient.MySqlConnection();
                conn.ConnectionString = myConnectionString;
                conn.Open();
                
                // Taking excel data and import them to the category and request model class
                categoryModel categoryModel = new categoryModel();
                requestModel requestModel = new requestModel();
                List<categoryModel> categoryresults = categoryModel.ReadFromCategory(@"C:.......\CategoryRecommender\CategoryRecommender\excelFolder\Category.xlsx");
                List<requestModel> requestresults = requestModel.ReadFromRequest(@"C:........\CategoryRecommender\CategoryRecommender\excelFolder\Request.xlsx");
                
                // Inserting data to the database
                foreach (categoryModel cmodel in categoryresults)
                {
                    var sql = "INSERT INTO testdb.category(id,title,tag) VALUES(" + cmodel.id + ",'" + cmodel.title + "','" + cmodel.tags + "')";
                    using (MySqlCommand commend = new MySqlCommand(sql, conn))
                    {
                        using (MySqlDataReader reader = commend.ExecuteReader())
                        {
                            commend.Parameters.AddWithValue("@cmodel.id", cmodel.id);
                            commend.Parameters.AddWithValue("@cmodel.title", cmodel.title);
                            commend.Parameters.AddWithValue("@cmodel.catId", cmodel.tags);
                        }
                    }
                } 
                foreach (requestModel rmodel in requestresults)
                {
                    var sql = "INSERT INTO testdb.request(id,title,catId,status,text) VALUES (@rmodel.id, @rmodel.title, @rmodel.catId, @rmodel.status,@rmodel.text);";
                    using (MySqlCommand commend = new MySqlCommand(sql, conn))
                    {
                        commend.Parameters.AddWithValue("@rmodel.id", rmodel.id);
                        commend.Parameters.AddWithValue("@rmodel.title", rmodel.title);
                        commend.Parameters.AddWithValue("@rmodel.catId", rmodel.catId);
                        commend.Parameters.AddWithValue("@rmodel.status", rmodel.status);
                        commend.Parameters.AddWithValue("@rmodel.text", rmodel.text);          
                        using (MySqlDataReader reader = commend.ExecuteReader())
                        {
                        }
                    }
                }
                MessageBox.Show("Data is successfully loaded.");
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // it takes title data from database, parses them into words and counts the same words. It finds most used words for each category. 
        private void button2_Click(object sender, EventArgs e)
        {


            string myConnectionString;
            myConnectionString = "server=localhost;uid=root;" + "pwd=1234;database=testdb";
            MySqlConnection myConnection = new MySqlConnection(myConnectionString);
            MySqlCommand myCommand = (MySqlCommand)myConnection.CreateCommand();
            myCommand.CommandText = "SELECT title,catId FROM testdb.request";
            myConnection.Open();       
            MySqlDataReader myReader = myCommand.ExecuteReader();

            // genelDict holds category id as a "key" and holds another dictionary called "wordDict" as a "value".
            // wordDict holds words as a "key" and counts of words as a "value".
            SortedDictionary<int, Dictionary<string, int>> genelDict = new SortedDictionary<int, Dictionary<string, int>>();
            
            // These are words that we don't want to count.
            string[] unusualWords = { "talebi", "sorunu","desteği", "hk.", "kurulumu", "için", "ve", "problemi", "hatası", "iç", "arızası", " ", "", ".", "-", "bilgisayarı", "bilgisayara", "işlemleri" ,
            "salonu","ile","diğer","hk","hakkında","sorun","talep","10","/","e","bilgisayarına","bilgisayarın","bilgisayarının","açamıyorum","yapamıyorum","mıyorum","iki","açılmıyor","a","adı",
            "çalışmıyor","e","bilgisayarım"};
            string[] possibleWords = {"kablo","açılması","ihtiyacı","değişimi","arıza","isteği"};     
            try
            {
                // Always call Read before accessing data.  
                while (myReader.Read())
                {      
                    var text = myReader.GetString(0).ToLower();
                    var content = Regex.Split(text, @" ");
                    Dictionary<string, int> wordDict = new Dictionary<string, int>();
                    if (text.Length > 0)
                    {
                        foreach (var word in content)
                        {
                            if (!( word.Length <3) && (!(word.ToString().Contains(@"+-/*"))))
                            {
                                if ((!unusualWords.Contains(word.ToString())) && (!word.Contains(@",\|!#$%&/()=?»«@£§€{}.;'<>_,")) && (!unusualWords.Contains(word.ToString())) && (!word.ToString().Any(char.IsDigit)) && (!(word.ToString().Any(char.IsPunctuation))))
                                {
                                    if (genelDict.ContainsKey(myReader.GetInt32(1)))
                                    {
                                        if (genelDict[myReader.GetInt32(1)].ContainsKey(word))
                                        {
                                            genelDict[myReader.GetInt32(1)][word] += 1;
                                        }
                                        else
                                        {
                                            //Dictionary<string, int> a = new Dictionary<string, int>();
                                            //a.Add(word, 1);
                                            genelDict[myReader.GetInt32(1)].Add(word, 1);
                                        }
                                    }
                                    else
                                    {
                                        //wordDict.Add(word, 1);
                                        Dictionary<string, int> a = new Dictionary<string, int>();
                                        a.Add(word, 1);
                                        genelDict.Add(myReader.GetInt32(1), a);
                                        //genelDict[myReader.GetInt32(1)].Add(word, 1);
                                    }
                                    wordDict.Clear();
                                }   
                            }
                        }
                    }
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // Close the connection when done with it.
                myConnection.Close();
            }
            MySql.Data.MySqlClient.MySqlConnection conn1;
            try
            {
                conn1 = new MySql.Data.MySqlClient.MySqlConnection();
                conn1.ConnectionString = myConnectionString;
                conn1.Open();
                
                // It inserts the nested dictionary to the database
                foreach (KeyValuePair<int, Dictionary<string, int>> kv in genelDict)
                {
                    //Debug.WriteLine(kv.Key.ToString() + " : " + kv.Value.ToArray());
                    foreach (KeyValuePair<string, int> ex in kv.Value)
                    {
                        //Debug.WriteLine(kv.Key.ToString() + "==> " + ex.Key.ToString() + " :" + ex.Value.ToString());
                        var sql1 = "INSERT INTO testdb.catCount(catId,tag,count) VALUES (@catID, @tag, @count);";      
                        if (ex.Value >= 5)
                        {
                            using (MySqlCommand commend1 = new MySqlCommand(sql1, conn1))
                            {
                                commend1.Parameters.AddWithValue("@catID", kv.Key);
                                commend1.Parameters.AddWithValue("@tag", ex.Key.ToString());
                                commend1.Parameters.AddWithValue("@count", ex.Value);
                                using (MySqlDataReader reader1 = commend1.ExecuteReader())
                                {
                                }
                            }
                        }
                    }
                }
                Thread.Sleep(2000);
                var join1 = "UPDATE testdb.catcount t1 INNER JOIN testdb.category t2 ON t1.catId = t2.id  SET t1.title = t2.title";
                using (MySqlCommand commend1 = new MySqlCommand(join1, conn1))
                {
                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                    {
                    }
                }

               


                MessageBox.Show("Data is successfully loaded.");
            }

            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        
        // It compares the old tag set with the most used words that we found on button 2 .
        private void button3_Click(object sender, EventArgs e)
        {
            string myConnectionString;
            myConnectionString = "server=localhost;uid=root;" + "pwd=1234;database=testdb";
            MySqlConnection myConnection = new MySqlConnection(myConnectionString);
            MySqlCommand myCommand = (MySqlCommand)myConnection.CreateCommand();
            myCommand.CommandText = "SELECT id,tag FROM testdb.category";
            myConnection.Open();
            MySqlDataReader myReader = myCommand.ExecuteReader();
            SortedDictionary<int, Dictionary<string, int>> genelDict = new SortedDictionary<int, Dictionary<string, int>>();
            MySql.Data.MySqlClient.MySqlConnection conn1;
            try
            {
                conn1 = new MySql.Data.MySqlClient.MySqlConnection();
                conn1.ConnectionString = myConnectionString;
                conn1.Open();
                // Always call Read before accessing data.
                while (myReader.Read())
                {
                    var text = myReader.GetString(1).ToLower();            
                    var content = Regex.Split(text, @" ");      
                    foreach (var word in content)
                    {
                        if (text.Length > 0)
                        {               
                            try
                            {                       
                                var sql1 = "INSERT INTO testdb.categorydb(id,tag) VALUES (@id, @tag);";
                                using (MySqlCommand commend1 = new MySqlCommand(sql1, conn1))
                                {
                                    commend1.Parameters.AddWithValue("@id", myReader.GetInt32(0));
                                    commend1.Parameters.AddWithValue("@tag",word);                        
                                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                                    {
                                    }
                                }             
                            }
                            catch (MySql.Data.MySqlClient.MySqlException ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }             
                    }
                }
                Thread.Sleep(2000);
                var join1 = "UPDATE testdb.categorydb t1 INNER JOIN testdb.category t2 ON t1.id = t2.id  SET t1.title = t2.title";
                using (MySqlCommand commend1 = new MySqlCommand(join1, conn1))
                {
                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                    {
                    }
                }
                Thread.Sleep(2000);
                var join2 = "UPDATE testdb.categorydb t1 INNER JOIN testdb.catcount t2 ON t1.tag = t2.tag  SET t1.count = t2.count";
                using (MySqlCommand commend1 = new MySqlCommand(join2, conn1))
                {
                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                    {
                    }
                }
                Thread.Sleep(2000);
                var deleteZeros = "DELETE FROM testdb.categorydb WHERE count = 0";
                using (MySqlCommand commend1 = new MySqlCommand(deleteZeros, conn1))
                {
                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                    {
                    }
                }
                



            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // Close the connection when done with it.
                myConnection.Close();
            }
            MessageBox.Show("Data is successfully loaded.");
        }
        
        // It is the same with button 2, but we parse text data instead of title data.
        private void button4_Click(object sender, EventArgs e)
        {
            string myConnectionString;
            myConnectionString = "server=localhost;uid=root;" + "pwd=1234;database=testdb";
            MySqlConnection myConnection = new MySqlConnection(myConnectionString);
            MySqlCommand myCommand = (MySqlCommand)myConnection.CreateCommand();
            myCommand.CommandText = "SELECT text,catId FROM testdb.request";
            myConnection.Open();
            MySqlDataReader myReader = myCommand.ExecuteReader();
            SortedDictionary<int, Dictionary<string, int>> genelDict = new SortedDictionary<int, Dictionary<string, int>>();
            string[] unusualWords = { "talebi", "sorunu","desteği", "hk.", "kurulumu", "için", "ve", "problemi", "hatası", "iç", "arızası", " ", "", ".", "-", "bilgisayarı", "bilgisayara", "işlemleri" ,
            "salonu","ile","diğer","hk","hakkında","sorun","talep","10","/","e","bilgisayarına","bilgisayarın","bilgisayarının","açamıyorum","yapamıyorum","mıyorum","iki","açılmıyor","a","adı",
            "çalışmıyor","e","bilgisayarım"};
            string[] possibleWords = { "kablo", "açılması", "ihtiyacı", "değişimi", "arıza", "isteği" };

            try
            {
                // Always call Read before accessing data.
                while (myReader.Read())
                {
                    var text = myReader.GetString(0).ToLower();
                    var content = Regex.Split(text, @" ");
                    if (text.Length > 0)
                    {
                        Dictionary<string, int> wordDict = new Dictionary<string, int>();
                        foreach (var word in content)
                        {
                            if (!(word.Length < 3) && (!(word.ToString().Contains(@"+-/*"))))
                            {
                                if ((!unusualWords.Contains(word.ToString())) && (!word.Contains(@",\|!#$%&/()=?»«@£§€{}.;'<>_,")) && (!unusualWords.Contains(word.ToString())) && (!word.ToString().Any(char.IsDigit)) && (!(word.ToString().Any(char.IsPunctuation))))
                                {
                                    if (genelDict.ContainsKey(myReader.GetInt32(1)))
                                    {
                                        if (genelDict[myReader.GetInt32(1)].ContainsKey(word))
                                        {
                                            genelDict[myReader.GetInt32(1)][word] += 1;
                                        }
                                        else
                                        {
                                            //Dictionary<string, int> a = new Dictionary<string, int>();
                                            //a.Add(word, 1);
                                            genelDict[myReader.GetInt32(1)].Add(word, 1);
                                        }
                                    }
                                    else
                                    {
                                        //wordDict.Add(word, 1);
                                        Dictionary<string, int> a = new Dictionary<string, int>();
                                        a.Add(word, 1);
                                        genelDict.Add(myReader.GetInt32(1), a);
                                        //genelDict[myReader.GetInt32(1)].Add(word, 1);
                                    }
                                    wordDict.Clear();
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // Close the connection when done with it.
                myConnection.Close();
            }
            MySql.Data.MySqlClient.MySqlConnection conn1;
            try
            {
                conn1 = new MySql.Data.MySqlClient.MySqlConnection();
                conn1.ConnectionString = myConnectionString;
                conn1.Open();
                foreach (KeyValuePair<int, Dictionary<string, int>> kv in genelDict)
                {
                    //Debug.WriteLine(kv.Key.ToString() + " : " + kv.Value.ToArray());
                    foreach (KeyValuePair<string, int> ex in kv.Value)
                    {
                        //Debug.WriteLine(kv.Key.ToString() + "==> " + ex.Key.ToString() + " :" + ex.Value.ToString());
                        var sql1 = "INSERT INTO testdb.textcount(catId,tag,count) VALUES (@catID, @tag, @count);";
                        //var sql = "INSERT INTO testdb.request(id,title,text,status) VALUES(" + @rmodel.id + ",'" + @rmodel.title + "','" + @rmodel.text + "','" + @rmodel.status + "')";
                        if (ex.Value >= 10)
                        {
                            using (MySqlCommand commend1 = new MySqlCommand(sql1, conn1))
                            {
                                commend1.Parameters.AddWithValue("@catID", kv.Key);
                                commend1.Parameters.AddWithValue("@tag", ex.Key.ToString());
                                commend1.Parameters.AddWithValue("@count", ex.Value);
                                using (MySqlDataReader reader1 = commend1.ExecuteReader())
                                {
                                }
                            }

                        }
                    }
                }
                Thread.Sleep(2000);
                var join1 = "UPDATE testdb.textcount t1 INNER JOIN testdb.category t2 ON t1.catId = t2.id  SET t1.title = t2.title";
                using (MySqlCommand commend1 = new MySqlCommand(join1, conn1))
                {
                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                    {
                    }
                    MessageBox.Show("Data is successfully loaded.");
                }

                
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // It is the same process on the button 3.
        private void button5_Click(object sender, EventArgs e)
        {
            string myConnectionString;
            myConnectionString = "server=localhost;uid=root;" + "pwd=1234;database=testdb";
            MySqlConnection myConnection = new MySqlConnection(myConnectionString);
            MySqlCommand myCommand = (MySqlCommand)myConnection.CreateCommand();
            myCommand.CommandText = "SELECT id,tag FROM testdb.category";
            myConnection.Open();
            MySqlDataReader myReader = myCommand.ExecuteReader();
            SortedDictionary<int, Dictionary<string, int>> genelDict = new SortedDictionary<int, Dictionary<string, int>>();
            MySql.Data.MySqlClient.MySqlConnection conn1;
            try
            {
                conn1 = new MySql.Data.MySqlClient.MySqlConnection();
                conn1.ConnectionString = myConnectionString;
                conn1.Open();
                // Always call Read before accessing data.
                while (myReader.Read())
                {
                    var text = myReader.GetString(1).ToLower();
                    var content = Regex.Split(text, @" ");
                    foreach (var word in content)
                    {
                        if (text.Length > 0)
                        {
                            try
                            {
                                var sql1 = "INSERT INTO testdb.textdb(id,tag) VALUES (@id, @tag);";
                                using (MySqlCommand commend1 = new MySqlCommand(sql1, conn1))
                                {
                                    commend1.Parameters.AddWithValue("@id", myReader.GetInt32(0));
                                    commend1.Parameters.AddWithValue("@tag", word);
                                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                                    {
                                    }
                                }
                            }
                            catch (MySql.Data.MySqlClient.MySqlException ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }
                    }
                }
                Thread.Sleep(2000);
                var join1 = "UPDATE testdb.textdb t1 INNER JOIN testdb.category t2 ON t1.id = t2.id  SET t1.title = t2.title";
                using (MySqlCommand commend1 = new MySqlCommand(join1, conn1))
                {
                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                    {
                    }
                }
                Thread.Sleep(2000);
                var join2 = "UPDATE testdb.textdb t1 INNER JOIN testdb.textcount t2 ON t1.tag = t2.tag  SET t1.count = t2.count";
                using (MySqlCommand commend1 = new MySqlCommand(join2, conn1))
                {
                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                    {
                    }
                }
                Thread.Sleep(2000);
                var deleteZeros = "DELETE FROM testdb.textdb WHERE count = 0";
                using (MySqlCommand commend1 = new MySqlCommand(deleteZeros, conn1))
                {
                    using (MySqlDataReader reader1 = commend1.ExecuteReader())
                    {
                    }
                }
                

            }
            finally
            {
                // always call Close when done reading.
                myReader.Close();
                // Close the connection when done with it.
                myConnection.Close();
            }
            MessageBox.Show("Data is successfully loaded.");
        }
       
        // It finds standard deviation of number of words.
        private void button6_Click(object sender, EventArgs e)
        {
            string myConnectionString;
            myConnectionString = "server=localhost;uid=root;" + "pwd=1234;database=testdb";
            MySqlConnection myConnection = new MySqlConnection(myConnectionString);
            MySqlCommand myCommand = (MySqlCommand)myConnection.CreateCommand();
            myCommand.CommandText = "SELECT tag,catId ,sum(count),STDDEV(count),MAX(count),AVG(count) from testdb.catcount GROUP by tag HAVING max(count) ORDER BY catId";
            myConnection.Open();
            MySqlDataReader myReader = myCommand.ExecuteReader();
            MySql.Data.MySqlClient.MySqlConnection conn1;
            try
            {
                conn1 = new MySql.Data.MySqlClient.MySqlConnection();
                conn1.ConnectionString = myConnectionString;
                conn1.Open();
                // Always call Read before accessing data.
                while (myReader.Read())
                {
                    var tag = myReader.GetString(0);
                    var catId = myReader.GetString(1);
                    var sum =  myReader.GetInt32(2);
                    var stdDev = Math.Round(myReader.GetDouble(3),2);
                    var max = Math.Round(myReader.GetDouble(4), 2);
                    var avg = Math.Round(myReader.GetDouble(5), 2);
                    try
                    {
                        var sql1 = "INSERT INTO testdb.deneyselbankacılık(tag,catId,sum,stdDev,max,avg) VALUES (@tag, @catId,@sum,@stdDev,@max,@avg);";
                        using (MySqlCommand commend1 = new MySqlCommand(sql1, conn1))
                        {
                            commend1.Parameters.AddWithValue("@tag",tag);
                            commend1.Parameters.AddWithValue("@catId", catId);
                            commend1.Parameters.AddWithValue("@sum", sum);
                            commend1.Parameters.AddWithValue("@stdDev", stdDev);
                            commend1.Parameters.AddWithValue("@max", max);
                            commend1.Parameters.AddWithValue("@avg", avg);
                            using (MySqlDataReader reader1 = commend1.ExecuteReader())
                            {
                            }
                        }
                    }
                    catch (MySql.Data.MySqlClient.MySqlException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
               
            }
            finally
            {

                // always call Close when done reading.
                myReader.Close();

                // Close the connection when done with it.
                myConnection.Close();
            }
            MessageBox.Show("Data is successfully loaded.");
        }
        
        // it exports data on the database to an excel file.
        private void button7_Click(object sender, EventArgs e)
        {
            ExportSql("testdb.catcount","title_count");
            MessageBox.Show("Excel file for 'Title Count' is created successfully.");
            Thread.Sleep(1000);

            ExportSql("testdb.categorydb", "title_match_count");
            MessageBox.Show("Excel file for 'Title Match' is created successfully.");
            Thread.Sleep(1000);


            ExportSql("testdb.textcount", "text_count");
            MessageBox.Show("Excel file for 'Text Count' is created successfully.");
            Thread.Sleep(1000);


            ExportSql("testdb.textdb", "text_match_count");
            MessageBox.Show("Excel file for 'Text Match' is created successfully.");
            Thread.Sleep(1000);

            ExportSql("testdb.deneyselbankacılık", "deneysel_bankacilik");
            MessageBox.Show("Excel file for 'Standart Deviation Calc. for Title' is created successfully.");

            MessageBox.Show("All excel files are created successfully!");


        }

        public static void ExportSql(string database, string excelFileName)
        {
            MySqlConnection cnn;
            string connectionString = null;
            string sql = null;
            string data = null;
            int i = 0;
            int j = 0;
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;
            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                connectionString = "server=localhost;uid=root;" + "pwd=1234;database=testdb";
                //conn1 = new MySqlConnection(connectionString);
                cnn = new MySqlConnection(connectionString);
                //conn1.Open();
                cnn.Open();
                MySqlDataAdapter dscmd;
                DataSet ds = new DataSet();
                switch (database)
                {
                    case "testdb.catcount":
                        var case1 = "SELECT 'id' as catId, 'tag' as tag,  'count' as count,'title' as title  UNION ALL select * from `catcount`;";
                        
                      
                        dscmd = new MySqlDataAdapter(case1, cnn);
                        dscmd.Fill(ds);
                        break;
                    case "testdb.categorydb":
                        var case2 = "SELECT 'id' as id, 'tag' as tag,  'title' as title,'count' as count  UNION ALL select * from `categorydb`;";

                        dscmd = new MySqlDataAdapter(case2, cnn);
                        dscmd.Fill(ds);


                        break;
                    case "testdb.textcount":
                        var case3 = "SELECT 'id' as catId, 'tag' as tag, 'count' as count, 'title' as title  UNION ALL select * from `textcount`;";

                        dscmd = new MySqlDataAdapter(case3, cnn);
                        dscmd.Fill(ds);
                        break;
                    case "testdb.textdb":
                        var case4 = "SELECT 'id' as id, 'tag' as tag,  'title' as title,'count' as count  UNION ALL select * from `textdb`;";

                        dscmd = new MySqlDataAdapter(case4, cnn);
                        dscmd.Fill(ds);
                        break;
                    case "testdb.deneyselbankacılık":
                        var case5 = "SELECT 'tag' as tag,'id' as catId, 'total' as sum,'Std_Dev' as stddev,'max' as max , 'avg' as avg  UNION ALL select * from `deneyselbankacılık`;";

                        dscmd = new MySqlDataAdapter(case5, cnn);
                        dscmd.Fill(ds);
                        break;
                    default:
                        break;

                }


                for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)

                {

                    for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)

                    {


                        data = ds.Tables[0].Rows[i].ItemArray[j].ToString();

                        xlWorkSheet.Cells[i + 1, j + 1] = data;

                    }

                }




                xlWorkBook.SaveAs(excelFileName+".xls",
                    Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                    Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);

                xlApp.Quit();


                ReleaseObject(xlWorkSheet);

                ReleaseObject(xlWorkBook);

                ReleaseObject(xlApp);

                MessageBox.Show("Excel file created , you can find the file in the 'My Documents' folder.");
            }
            catch (Exception xmessage)
            {
                MessageBox.Show(xmessage.ToString());
            }

            static void ReleaseObject(object obj)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;


                }

                catch (Exception ex)
                {
                    obj = null;
                    MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
                }

                finally

                {

                    GC.Collect();

                }




            }
        }

        
    }
}

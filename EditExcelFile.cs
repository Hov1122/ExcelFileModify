using System;
using System.IO;
using System.Data.SQLite;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Deployment.Application;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelFileModify
{
    internal class EditExcelFile
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public static List<Process> childrenProcess = new List<Process>(); // keep track of children process to destroy them on app force exit

        public static SQLiteConnection CreateConnection()
        {

            SQLiteConnection sqlite_conn;
            // Create a new database connection:
            string dataPath = ApplicationDeployment.CurrentDeployment.DataDirectory + @"\Database";
            //string dataPath = Environment.CurrentDirectory + @"\Database";

            bool exists = !(File.Exists(dataPath + "\\database.db"));

            if (!exists)
                System.Threading.Thread.Sleep(1000);
            sqlite_conn = new SQLiteConnection("Data Source=" + dataPath +
                "\\database.db; Version = 3; New = True; Compress = True; ");
            // Open the connection:
            
            try
            {
                sqlite_conn.Open();
            }
            catch (Exception)
            {
                throw new SQLiteException("cant connect to the database.");
            }
            if (!exists)
            {
                CreateTable(sqlite_conn);
            }
            return sqlite_conn;
        }

        static void CreateTable(SQLiteConnection conn)
        {

            SQLiteCommand sqlite_cmd;
            string Createsql = "CREATE TABLE IF NOT EXISTS Zones " +
               "(id INTEGER PRIMARY KEY, zone TINYINT, street TEXT UNIQUE)";

            sqlite_cmd = conn.CreateCommand();
            sqlite_cmd.CommandText = Createsql;
            sqlite_cmd.ExecuteNonQuery();

            string add_indexes = "CREATE INDEX IF NOT EXISTS street_index ON Zones (street);";
            sqlite_cmd.CommandText = add_indexes;
            sqlite_cmd.ExecuteNonQuery();

        }

        public static Task ClearData()
        {
            SQLiteCommand sqlite_cmd;
            SQLiteConnection conn;
            try
            {
                conn = CreateConnection();
            }
            catch
            {
                throw new SQLiteException("cant connect to the database.");
            }

            sqlite_cmd = conn.CreateCommand();
            sqlite_cmd.CommandText = "SELECT COUNT(*) FROM Zones";

            SQLiteDataReader sqlite_datareader = sqlite_cmd.ExecuteReader();

            if (sqlite_datareader.Read() && sqlite_datareader.GetInt16(0) == 0)
            {
                sqlite_datareader.Close();
                sqlite_datareader.Dispose();
                conn.Close();
                throw new MissingFieldException("No data to clear");
            }

            sqlite_datareader.Close();
            sqlite_datareader.Dispose();
            
            string deletesql = "DELETE FROM Zones;" +
               "VACUUM;";
            sqlite_cmd = conn.CreateCommand();
            sqlite_cmd.CommandText = deletesql;
            sqlite_cmd.ExecuteNonQuery();

            conn.Close();

            return Task.CompletedTask;

        }

        static void InsertData(SQLiteConnection conn, int z, string s)
        {
            SQLiteCommand sqlite_cmd;
            sqlite_cmd = conn.CreateCommand();
            sqlite_cmd.CommandText = "Insert INTO Zones " +
               "(zone, street) VALUES(@zone, @street) " +
               "ON CONFLICT(street) DO UPDATE " +
               "SET street=@street, zone=@zone ;";
            sqlite_cmd.Parameters.Add(new SQLiteParameter("@zone", z));
            sqlite_cmd.Parameters.Add(new SQLiteParameter("@street", s));
            sqlite_cmd.ExecuteNonQuery();


        }
        
        public static Task ImportExcelData(string path)
        {
            if (path.Length == 0) return Task.CompletedTask;

            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                throw new ArgumentNullException("couldnt find Excel");
            }
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); ;

            int id;
            GetWindowThreadProcessId(xlApp.Hwnd, out id);
            childrenProcess.Add(Process.GetProcessById(id));

            //Console.OutputEncoding = Encoding.UTF8;
            // Console.WriteLine(xlWorkSheet.Cells[6, 3].Text);
            string a;
            int street_col = 0;
            int zone_col = xlWorkSheet.UsedRange.Columns.Count;
            bool check_zone = false, check_street = false;
            for (int i = 1; i <= xlWorkSheet.UsedRange.Columns.Count; i++)
            {
                a = xlWorkSheet.Cells[1, i].Text;
                if (String.Equals(a, "улица", StringComparison.OrdinalIgnoreCase))
                {
                    street_col = i;
                    check_street = true;
                }
                else if (String.Equals(a, "Зоны доставки", StringComparison.OrdinalIgnoreCase))
                {
                    zone_col = i;
                    check_zone = true;
                }
            }

            if (!(check_zone && check_street))
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlApp);
                throw new ArgumentException("1 toxum piti lini улица ev Зоны доставки syunakner");
            }

            SQLiteConnection sqlite_conn;
            try
            {
                sqlite_conn = CreateConnection();
            }
            catch
            {
                xlWorkBook.Close(true);
                xlApp.Quit();
                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlApp);
                throw new SQLiteException("cant connect to the database.");
            }

            for (int i = 1; i <= xlWorkSheet.UsedRange.Rows.Count; i++)
            {

                try
                {
                    int z = (int)Double.Parse(xlWorkSheet.Cells[i, zone_col].Text);
                    //Console.OutputEncoding = Encoding.UTF8;

                    if (xlWorkSheet.Cells[i, street_col].Text.Length != 0)
                    {
                        string tmp = xlWorkSheet.Cells[i, street_col].Text;
                        tmp = tmp.Trim();
                        InsertData(sqlite_conn, z, tmp.ToLower());
                    }


                }
                catch (FormatException)
                {
                    // do nothing
                }

            }

            //ReadData(sqlite_conn);
            xlWorkBook.Close(true);
            xlApp.Quit();
            ReleaseObject(xlWorkSheet); 
            ReleaseObject(xlWorkBook);
            Process proc = Process.GetProcessById(id);
            proc.Kill();
            //ReleaseObject(xlApp);

            return Task.CompletedTask;
        }

        public static  Task<List<string>> AddZones(string path, bool in_parts) //async
        {
            if (path.Length == 0) throw new ArgumentException();

            List<string> tmp_paths = new List<string>();

            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                throw new ArgumentNullException("couldnt find excel");
            }

            string dataPath = ApplicationDeployment.CurrentDeployment.DataDirectory;
            //string dataPath = Environment.CurrentDirectory;
            int id;
            GetWindowThreadProcessId(xlApp.Hwnd, out id);
            Process proc = Process.GetProcessById(id);
            childrenProcess.Add(proc);

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            xlApp.DisplayAlerts = false;

            SQLiteDataReader sqlite_datareader;
            SQLiteCommand sqlite_cmd;
            SQLiteConnection conn;

            try
            {
                conn = CreateConnection();
            }
            catch
            {
                throw new SQLiteException("cant open db");
            }

            sqlite_cmd = conn.CreateCommand();
            sqlite_cmd.CommandText = "SELECT COUNT(*) FROM Zones";

            sqlite_datareader = sqlite_cmd.ExecuteReader();

            if (sqlite_datareader.Read() && sqlite_datareader.GetInt16(0) == 0)
            {
                sqlite_datareader.Close();
                sqlite_datareader.Dispose();
                conn.Close();
                xlWorkBook.Close(true);
                xlApp.Quit();
                ReleaseObject(xlWorkBook);
                proc.Kill();
               // ReleaseObject(xlApp);
                throw new MissingFieldException("Import some data first");
            }

            sqlite_datareader.Close();
            sqlite_datareader.Dispose();

            string missingStreetFilePath = dataPath + @"\tmpFiles\" + "missing.txt";

            if (File.Exists(missingStreetFilePath))
                File.Delete(missingStreetFilePath);

            Dictionary<string, Excel.Range> Couriers = new Dictionary<string, Excel.Range>();

            Excel.Sheets tmpWorkSheets = xlWorkBook.Worksheets;

            int total = 0;
            int totalCourier = 0;
            
            foreach (Excel.Worksheet xlWorkSheet in tmpWorkSheets)
            {

                Excel.Range ur = xlWorkSheet.UsedRange;
                Excel.Range r = xlWorkSheet.Cells[1, ur.Columns.Count];

                if (xlWorkSheet.Cells[1, r.Column].Text == "")
                    r = r.get_End(Excel.XlDirection.xlToLeft);

                int zone_col = r.Column + 1;
                int courier_col = 3;

                bool checkZone = false, checkCourier = false;
               
                for (int i = 1; i <= xlWorkSheet.UsedRange.Columns.Count; i++)
                {
                    string a = xlWorkSheet.Cells[1, i].Text;
                    a = a.Trim();
                    if (String.Equals(a, "Зоны доставки", StringComparison.OrdinalIgnoreCase))
                    {
                        zone_col = i;
                        checkZone = true;
                    }
                    else if (String.Equals(a, "Курьер", StringComparison.OrdinalIgnoreCase))
                    {
                        courier_col = i;
                        checkCourier = true;
                    }
                    if (checkCourier && checkZone)
                        break;
                }

                xlWorkSheet.Cells[1, zone_col].Value = "Зоны доставки";
                xlWorkSheet.Cells[1, zone_col + 1].Value = "Сумма";

                int street_col = 1;

                bool check_street = false;
                for (int i = 1; i <= xlWorkSheet.UsedRange.Columns.Count; i++)
                {
                    string a = xlWorkSheet.Cells[1, i].Text;
                    a = a.Trim();
                    if (String.Equals(a, "улица", StringComparison.OrdinalIgnoreCase))
                    {
                        street_col = i;
                        check_street = true;
                        break;
                    }
                }

                if (!check_street || (!checkCourier && in_parts))
                {
                    conn.Close();
                    xlWorkBook.Saved = true;
                    xlWorkBook.Close(true);
                    xlApp.Quit();
                    ReleaseObject(xlWorkSheet);
                    ReleaseObject(xlWorkBook);
                    proc.Kill();
                    //ReleaseObject(xlApp);
                    if (in_parts)
                        throw new ArgumentException("1 toxum piti lini улица ev Курьер");
                    else
                        throw new ArgumentException("1 toxum piti lini улица");
                }

                sqlite_cmd.CommandText = "SELECT [zone] FROM Zones WHERE [street] = @s";

                int courier_start = 1;
                for (int i = 2; i <= xlWorkSheet.UsedRange.Rows.Count; i++)
                {
                    if (xlWorkSheet.Cells[i, courier_col].Text.ToString() != "")
                    {
                        courier_start = i;
                        break;
                    }
                }
                
                int courier_end = courier_start;
                string courier_name = xlWorkSheet.Cells[2, courier_col].Text;

                Excel.Range c1;
                Excel.Range c2;
                Excel.Range rng;

                HashSet<string> missingStreets = new HashSet<string>();

                for (int i = 2; i <= xlWorkSheet.UsedRange.Rows.Count; i++)
                {
                    
                    string st = xlWorkSheet.Cells[i, street_col].Text;
                    st = st.Trim();
                    if (st.Length > 0)
                    {
                        sqlite_cmd.Parameters.Add(new SQLiteParameter("@s", st.ToLower()));
                        sqlite_datareader = sqlite_cmd.ExecuteReader();

                        if (sqlite_datareader.Read())
                        {
                            int zone = sqlite_datareader.GetInt16(0);
                            xlWorkSheet.Cells[i, zone_col].Value = zone;
                            if (zone == 1) xlWorkSheet.Cells[i, zone_col + 1].Value = 150;
                            else xlWorkSheet.Cells[i, zone_col + 1].Value = zone * 100;
                            
                            totalCourier += xlWorkSheet.Cells[i, zone_col + 1].Value;
                        }
                        else
                        {            
                            using (StreamWriter writer = new StreamWriter(missingStreetFilePath, true))
                            {
                                if (!missingStreets.Contains(st))
                                {
                                    writer.WriteLine($"{missingStreets.Count + 1}. {st}");
                                    missingStreets.Add(st);
                                }
                            } 
                        }
                        sqlite_datareader.Close();
                        sqlite_datareader.Dispose();
                    }

                    if (xlWorkSheet.Cells[courier_end, courier_col].Text.Length > 0 && courier_name != xlWorkSheet.Cells[courier_end, courier_col].Text.ToString())
                    {
                        c1 = xlWorkSheet.Cells[courier_start, 1];
                        c2 = xlWorkSheet.Cells[courier_end, xlWorkSheet.UsedRange.Columns.Count];

                        if (courier_end - courier_start > 1)
                        {
                            rng = xlWorkSheet.Range[c1, c2];
                            try
                            {
                                Couriers.Add(courier_name, rng);
                                if (xlWorkSheet.Cells[i, courier_col].Text.ToString().Contains(courier_name))
                                {
                                    xlWorkSheet.Cells[i, zone_col + 1] = totalCourier;
                                    total += totalCourier;
                                    totalCourier = 0;
                                    xlWorkSheet.Cells[i, zone_col + 1].NumberFormat = "###,#";
                                }
                            }
                            
                            catch (ArgumentException)
                            {
                                // do nothing
                            }
                        }
                        
                        courier_start = courier_end;        
                        courier_name = xlWorkSheet.Cells[courier_end, courier_col].Text;

                    }
                    courier_end++;          
                }
            
                xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, zone_col + 1].Value = total;
                xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, zone_col + 1].NumberFormat = "###,#";

                xlWorkSheet.UsedRange.Columns.AutoFit();
                xlWorkSheet.UsedRange.Rows.AutoFit();

                if (in_parts)
                {
                    if (xlWorkBook.Worksheets.Count == 1)
                    {
                        foreach(var courier in Couriers)
                        {
                            Excel.Worksheet courierWSh;
                            courierWSh = (Excel.Worksheet)xlWorkBook.Worksheets.Add(xlWorkBook.Worksheets.get_Item(xlWorkBook.Worksheets.Count));
                            courierWSh.Name = courier.Key;

                            string[] tmpRange = courier.Value.Address.Split('$');
                            Excel.Range from = xlWorkSheet.Range["A1:" + tmpRange[3][0] + "1"];                
                            Excel.Range to = courierWSh.Range["A1:" + tmpRange[3][0] + "1"];

                            from.Copy(to);

                            int rngHeight = 1 + Int16.Parse(tmpRange[4]) - Int16.Parse(tmpRange[2].Substring(0, tmpRange[2].Length - 1));

                            from = xlWorkSheet.Range[courier.Value.Address];
                            to = courierWSh.Range["A2:" + tmpRange[3][0] + rngHeight];
                            
                            from.Copy(to);

                            if (courierWSh.Cells[2, 1].Text.ToString() == "" && courierWSh.Cells[2, 2].Text.ToString() == "")
                            {
                                from = xlWorkSheet.Range["A2:B2"];
                                to = courierWSh.Range["A2:B2"];
                                from.Copy(to);
                            }

                            courierWSh.Cells[courierWSh.UsedRange.Rows.Count, zone_col + 1].NumberFormat = "###,#";

                            courierWSh.UsedRange.Columns.AutoFit();
                            courierWSh.UsedRange.Rows.AutoFit();
                            xlWorkBook.SaveAs(dataPath + @"\tmpFiles\" + xlWorkBook.Name);
                            tmp_paths.Add(dataPath + @"\tmpFiles\" + xlWorkBook.Name);
                            xlWorkBook.Saved = true;
                            ReleaseObject(courierWSh);
                        }
                    }
                    else
                    {
                        Excel.Workbook xlWorkBookPart = xlApp.Workbooks.Add();
                        xlApp.SheetsInNewWorkbook = 1;
                        Excel.Worksheet xlWorkSheetPart = (Excel.Worksheet)xlWorkBookPart.Worksheets.get_Item(1);
                        xlWorkSheet.UsedRange.Copy();
                        xlWorkSheetPart.Paste();
                        xlWorkSheetPart.Name = xlWorkSheet.Name;
                        xlWorkSheetPart.UsedRange.Columns.AutoFit();
                        xlWorkSheetPart.UsedRange.Rows.AutoFit();

                        int li = xlWorkBook.Name.LastIndexOf(".xlsx");
                        if (li != -1 && xlWorkSheet.Name == xlWorkBook.Name.Substring(0, li))
                        {
                            xlWorkSheetPart.SaveAs(dataPath + @"\tmpFiles\" + "tt" + xlWorkSheet.Name);
                            tmp_paths.Add(dataPath + @"\tmpFiles\" + "tt" + xlWorkSheet.Name + ".xlsx");
                        }
                            
                        else
                        {
                            xlWorkSheetPart.SaveAs(dataPath + @"\tmpFiles\" + xlWorkSheet.Name);
                            tmp_paths.Add(dataPath + @"\tmpFiles\" + xlWorkSheet.Name + ".xlsx");
                        }
                            
                        xlWorkBookPart.Saved = true;
                        xlWorkBookPart.Close(true);
                        ReleaseObject(xlWorkSheetPart);
                        ReleaseObject(xlWorkBookPart);
                    }
                    
                }
                    
                ReleaseObject(xlWorkSheet);

            }
            conn.Close();
            
            if (!in_parts)
            {
                string tmpFilePath = dataPath + @"\tmpFiles\" + xlWorkBook.Name;
                xlWorkBook.SaveAs(tmpFilePath);
                tmp_paths.Clear();
                tmp_paths.Add(dataPath + @"\tmpFiles\"  + xlWorkBook.Name);
            }
            
            xlWorkBook.Saved = true;
            xlWorkBook.Close(true);
            xlApp.Quit();
            ReleaseObject(xlWorkBook);
            proc.Kill();
            //ReleaseObject(xlApp);

            if (File.Exists(missingStreetFilePath))
                tmp_paths.Add(missingStreetFilePath);
            return Task.FromResult(tmp_paths);
           // return tmp_paths;

        }
        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}

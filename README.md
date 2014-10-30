Excel_NPOI
==========
 protected void Page_Load(object sender, EventArgs e)
    {
 
    }
    protected void Unnamed1_Click(object sender, EventArgs e)
    {
        string path = @"C:\Users\OI\Documents\test.xls";
 
        //甲資料
        DataTable dt = new DataTable();
        dt.Columns.Add("aaa");
        dt.Columns.Add("bbb");
        dt.Columns.Add("ccc");
        dt.Columns.Add("ddd");
        dt.Rows.Add(dt.NewRow());
        dt.Rows.Add(dt.NewRow());
        dt.Rows.Add(dt.NewRow());
        dt.Rows[0]["aaa"] = "test";
        dt.Rows[1]["aaa"] = "test";
        dt.Rows[2]["aaa"] = "test";
        dt.Rows[0]["bbb"] = "test";
        dt.Rows[1]["bbb"] = "test";
        dt.Rows[2]["ccc"] = "test";
 
        //設定資料庫 , 匯入模版, 要空的行數
        DataTableToExcelFile(dt, path, 3);
 
    }
 
    //DataTable轉成Excel檔案的方法
    private void DataTableToExcelFile(DataTable dt, string templetPath, int x = 5)
    {
        //匯入模版
        FileStream files = new FileStream(templetPath, FileMode.Open, FileAccess.Read);
 
        //建立Excel 2003檔案
        IWorkbook wb = new HSSFWorkbook(files);
        ISheet ws;
 
        ////建立Excel 2007檔案
        //IWorkbook wb = new XSSFWorkbook();
        //ISheet ws;
       
 
        if (dt.TableName != string.Empty)
        {
            ws = wb.CreateSheet(dt.TableName);
        }
        else
        {
            ws = wb.GetSheetAt(0);
        }
 
        ws.CreateRow(x);//第x行為欄位名稱
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            ws.GetRow(x).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
        }
 
 
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            ws.CreateRow(i + 1 + x);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ws.GetRow(i + 1 + x).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
            }
        }
 
        FileStream file = new FileStream(@"D:\npoi.xls", FileMode.Create);//產生檔案的位置
        wb.Write(file);
        file.Close();
    }

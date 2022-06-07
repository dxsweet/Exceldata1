using System.Diagnostics;
using System.Text;
using Excel1 = Microsoft.Office.Interop.Excel;

internal class Program
{
    static void Main(String[] args)
    {

        foreach(Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("EXCEL"))
            {
                clsProcess.Kill();
            }
        }

        do
        {
            Console.WriteLine("0:退出程序");
            Console.WriteLine("1:跑道视程复制到跑道视程1");
            Console.WriteLine("2:从ds文件夹提取数据到output");
            Console.WriteLine("请输入数字进行下一步");
            String option = Console.ReadLine();

            if(option == "0")
            {
                Console.WriteLine("退出程序");
                break;
            }
            else if(option == "1")
            {
                Program.HandleExcel1();
            }
            else if(option == "2")
            {
                Program.HandleExcel2();
            }
            else
            {
                Console.WriteLine("输入无效,请重新输入");
            }

        } while (true);

    }




    static void HandleExcel1()
    {
        Excel1.Application xlApp = new Excel1.Application();
        xlApp.Visible = true;
        xlApp.DisplayAlerts = false;

        DirectoryInfo di = new DirectoryInfo(@"./sd");

        //把跑到视程2和3复制到1.
        foreach (FileInfo fi in di.GetFiles())
        {
            Excel1.Workbook wb = xlApp.Workbooks.Open(fi.FullName);
            Excel1.Worksheet ws1 = wb.Worksheets["跑道视程一"];
            Excel1.Worksheet ws2 = wb.Worksheets["跑道视程二"];
            Excel1.Worksheet ws3 = wb.Worksheets["跑道视程三"];

            Excel1.Range ws1r1 = ws1.Range["a48", "z91"];
            Excel1.Range ws1r2 = ws1.Range["a92", "z127"];

            Excel1.Range ws2r = ws2.Range["a4", "y47"];
            ws2r.Copy(ws1r1);

            Excel1.Range ws3r = ws3.Range["a4", "y39"];
            ws3r.Copy(ws1r2);

            wb.SaveAs2(fi.FullName);

            wb.Close();
            xlApp.Workbooks.Close();

        }


        xlApp.Quit();

    }

    static void HandleExcel2()
    {
        Excel1.Application xlApp = new Excel1.Application();
        xlApp.Visible = false;
        xlApp.DisplayAlerts = false;


        String[] allLines = File.ReadAllLines(@"./td1.txt");

        DateTime dt0, dt1, dt0utc, dt1utc;
        

        StringBuilder sb1 = new StringBuilder();
        int i = 1;
        foreach (String line in allLines)
        {

            String[] ts2s = line.Trim().Split('-');
            dt0 = new DateTime(Convert.ToInt32(ts2s[0]), Convert.ToInt32(ts2s[1]), Convert.ToInt32(ts2s[2]), Convert.ToInt32(ts2s[3]), 0, 0);
            dt1 = new DateTime(Convert.ToInt32(ts2s[4]), Convert.ToInt32(ts2s[5]), Convert.ToInt32(ts2s[6]), Convert.ToInt32(ts2s[7]), 0, 0);

            //Console.WriteLine(dt0);
            //Console.WriteLine(dt1);
            dt0utc = dt0.AddHours(-8);
            dt1utc = dt1.AddHours(-8);


            Excel1.Workbook wb = xlApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + "\\sd\\" + ts2s[0] + ts2s[1] + ".xls");
            Excel1.Worksheet ws1 = wb.Worksheets["温度"];
            Excel1.Worksheet ws2 = wb.Worksheets["露点温度"];
            Excel1.Worksheet ws3 = wb.Worksheets["相对湿度"];
            Excel1.Worksheet ws4 = wb.Worksheets["修正海平面气压"];
            Excel1.Worksheet ws5 = wb.Worksheets["风向风速"];
            Excel1.Worksheet ws7 = wb.Worksheets["主导能见度"];
            Excel1.Worksheet ws8 = wb.Worksheets["跑道视程一"];



            //Cells.Item(Row, Column)
            //Value2 属性和 Value 属性的唯一区别在于 Value2 属性不使用 Currency 和 Date 数据类型
            
            int j = 1;
            for (DateTime dtx = dt0; dtx <= dt1; dtx = dtx.AddHours(1))
            {
                sb1.Append(dtx.ToString("yyyyMMddHH") + ",");
                DateTime dtxutc = dtx.AddHours(-8);
                sb1.Append(dtxutc.ToString("yyyyMMddHH") + ",");
                //1温度
                int ws1row = 1;
                int ws1col = 1;

                int dtxd = Int32.Parse(dtx.ToString("dd"));
                int dtxh = Int32.Parse(dtx.ToString("HH"));

                if (dtxh == 0)
                {
                    if(dtxd == 1)
                    {
                        Console.WriteLine("北京时间:"+dtx.ToString("yyyyMMddHH") + "  UTC时间:"+ dtxutc.ToString("yyyyMMddHH") + "  没有数据");
                        continue;
                    }

                    dtxh = 24;
                    dtxd = dtxd - 1;
                }



                if (dtxd <= 10)
                {
                    ws1row = dtxd + 3;

                }
                else if (dtxd <= 20)
                {
                    ws1row = dtxd + 5;
                }
                else if (dtxd <= 31)
                {
                    ws1row = dtxd + 7;
                }

                ws1col = dtxh + 1;



                if ((ws1.Cells[ws1row, ws1col].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws1.Cells[ws1row, ws1col].value2.ToString() + ",");
                }


                //Console.WriteLine("行号:"+ ws1row + ", 列号:" + ws1col + ", 结果为:" + ws1.Cells[ws1row, ws1col].value2.ToString());
                //Console.ReadLine();

                //2露点温度
                int ws2row = dtxd + 3;

                if ((ws2.Cells[ws2row, ws1col].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws2.Cells[ws2row, ws1col].value2.ToString() + ",");
                }

                //3相对湿度



                if ((ws3.Cells[ws1row, ws1col].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws3.Cells[ws1row, ws1col].value2.ToString() + ",");
                }



                //4修正海压

                if ((ws4.Cells[ws1row, ws1col].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws4.Cells[ws1row, ws1col].value2.ToString() + ",");
                }


                //5风向


                int ws5col = dtxh * 2;


                if ((ws5.Cells[(ws1row + 2), ws5col].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws5.Cells[ws1row + 2, ws5col].value2.ToString() + ",");
                }


                //6风速

                if ((ws5.Cells[ws1row + 2, ws5col + 1].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws5.Cells[ws1row + 2, ws5col + 1].value2.ToString() + ",");
                }



                //7主导能见度


                if ((ws7.Cells[ws2row, ws1col].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws7.Cells[ws2row, ws1col].value2.ToString() + ",");
                }


                //8跑道视程2列
                int ws8row = dtxd * 4;


                if ((ws8.Cells[ws8row, ws1col].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws8.Cells[ws8row, ws1col].value2.ToString() + ",");
                }
              //  Console.WriteLine(ws8row + ":" + ws1col + "值 " + ws8.Cells[ws8row, ws1col].value2.ToString());

                if ((ws8.Cells[ws8row + 1, ws1col].value2) == null)
                {
                    sb1.Append("null,");
                }
                else
                {
                    sb1.Append(ws8.Cells[ws8row + 1, ws1col].value2.ToString() + ",");
                }


                sb1.Append(j + ",\n");


                if (j == 1)
                {
                    sb1.Remove(sb1.Length - 1, 1);
                    sb1.Append(i + ":" + dt0utc.ToString("yyyy年MM月dd日HH时UTC到") + dt1utc.ToString("yyyy年MM月dd日HH时UTC") + "\n");
                }

                j++;

               
            }





            i++;





            wb.Close();
            xlApp.Workbooks.Close();
            Console.WriteLine(line + "已经处理完毕");

        }


        String text1 = sb1.ToString().Trim();
        System.IO.File.WriteAllText(@"./output.txt", text1, System.Text.Encoding.UTF8);


        xlApp.Quit();
        Console.WriteLine("全部处理完毕");

        foreach (Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("EXCEL"))
            {
                clsProcess.Kill();
            }
        }


        Console.ReadLine();
    }
}
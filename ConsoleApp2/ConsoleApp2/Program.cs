using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeOpenXml;
struct ver
{
    public long vertices;
    public double wight;
};
struct ver2
{
    public long s1;
    public long s2;
};
struct ver3
{
    public string f1;
    public string f2;
    public string hyperlink1;
    public string hyperlink2;
    public double lineMatch;
};
public class solution
{
    public static List<double> l5 = new List<double>();//store avg of all component /used in stat file
    static Dictionary<long, List<ver>> d = new Dictionary<long, List<ver>>();//adjlist
    static Dictionary<double, List<ver2>> d2 = new Dictionary<double, List<ver2>>();//take its wight and return all edge with this number /used in mst file
    static Dictionary<double, List<ver3>> d3 = new Dictionary<double, List<ver3>>();//used in sorting /used in mst file
    static Dictionary<double, long> indexOfConnComp = new Dictionary<double, long>();//used to return the id index of the connected component and its key is avg of the component /used in stat file
    static List<long> vertice = new List<long>();//vertices used in stat file
    public static List<long>[] l6;//store all component /used in mst file
    static Dictionary<long, bool> visited2 = new Dictionary<long, bool>();//check if vertices visited or not /used in stat file
    static Dictionary<long, long> id = new Dictionary<long, long>();//take vertices and return its id /used in mst file
    static Dictionary<long, List<long>> visited4 = new Dictionary<long, List<long>>();//take id and return list of the vertices /used in mst file
    public static List<double> list1 = new List<double>();//used in sorting /used in mst file
    public static List<double> list2 = new List<double>();//used in sorting /used in mst file
    static List<ver2> list3 = new List<ver2>();
    static ver v = new ver();
    static ver2 v2 = new ver2();
    static ver3 v3 = new ver3();
    public static void WriteDataToExcelWithEpplus(string filePath, long x, List<long>[] l6, double[] maxw) //o(v)
    {
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            worksheet.Cells[1, 1].Value = "Component Index";
            worksheet.Cells[1, 2].Value = "Vertices";
            worksheet.Cells[1, 3].Value = "Average Similarity";
            worksheet.Cells[1, 4].Value = "Component Count";
            l5.Sort((x, y) => y.CompareTo(x));//o(Glog G)
            int k = 0;//no of component
            //l5 is maxw descending order
            foreach (var i in l5)
            {
                
                string combine = "";
                //to show vertices ordered in each component
                l6[indexOfConnComp[i]].Sort();//(v log v)
                //loop through component vertices
                Console.WriteLine(i);
                foreach (var v in l6[indexOfConnComp[i]])
                {
                    Console.WriteLine(v);
                    combine += v;
                    //so no , at the end
                    if (v != l6[indexOfConnComp[i]][l6[indexOfConnComp[i]].Count - 1])
                    {
                        combine += ", ";
                    }
                }
                //put values in columns
                worksheet.Cells[k + 2, 3].Value = Math.Round(maxw[indexOfConnComp[i]], 1);
                worksheet.Cells[k + 2, 2].Value = combine;
                worksheet.Cells[k + 2, 1].Value = k + 1;
                worksheet.Cells[k + 2, 4].Value = l6[indexOfConnComp[i]].Count;
                k++;
            }
            worksheet.Cells.AutoFitColumns();
            //save excel file in that path
            package.SaveAs(filePath);
        }
    }
    public static (List<string>, List<string>, List<double>) ReadDataToLists(string filename)
    {
        List<string> col1 = new List<string>();
        List<string> col2 = new List<string>();
        List<double> col3 = new List<double>();
        long count = 0;
        double frac = 0;
        using (ExcelPackage package = new ExcelPackage(new FileInfo(filename))) //o(n)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
            for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
            {
                if (worksheet.Cells[row, 3].Value == null)
                    continue;
                col1.Add((string)worksheet.Cells[row, 1].Value);
                col2.Add((string)worksheet.Cells[row, 2].Value);
                string l = (string)worksheet.Cells[row, 1].Value;
                string l2 = (string)worksheet.Cells[row, 2].Value;
                col3.Add((double)worksheet.Cells[row, 3].Value);
                long extractedpath = ext(l);
                long extractedpath2 = ext(l2);
                double p1 = extper(l);
                double p2 = extper(l2);
                double line = (double)worksheet.Cells[row, 3].Value;
                if (!d.ContainsKey(extractedpath))
                {
                    d[extractedpath] = new List<ver>();
                    vertice.Add(extractedpath);
                    visited2[extractedpath] = false;
                    count++;
                    visited4[count] = new List<long>();
                    id[extractedpath] = count;
                    visited4[count].Add(extractedpath);
                }
                if (!d.ContainsKey(extractedpath2))
                {
                    d[extractedpath2] = new List<ver>();
                    vertice.Add(extractedpath2);
                    visited2[extractedpath2] = false;
                    count++;
                    visited4[count] = new List<long>();
                    id[extractedpath2] = count;
                    visited4[count].Add(extractedpath2);
                }
                v2.s1 = extractedpath;
                v2.s2 = extractedpath2;
                if (list1.Contains(Math.Max(p1, p2) + frac + line / 5000))
                    frac -= 0.0000000000001;
                d2[Math.Max(p1, p2) + frac + line / 5000] = new List<ver2>();
                d2[Math.Max(p1, p2) + frac + line / 5000].Add(v2);
                list1.Add(Math.Max(p1, p2) + frac + line / 5000);
                v.wight = (p1 + p2) / 2;
                v.vertices = extractedpath2;
                d[extractedpath].Add(v);
                v.vertices = extractedpath;
                d[extractedpath2].Add(v);
            }
            worksheet.Cells.AutoFitColumns();
            package.Save();
        }
        return (col1, col2, col3);
    }
    public static double extper(string item) //o(1)
    {
        string pattern = @"\d+(?=\%)";
        Match match = Regex.Match(item, pattern);
        double number = 0;
        if (match.Success)
        {
            string numberString = match.Value;
            number = long.Parse(numberString);
        }
        return number;
    }
    public static long ext(string path)  //o(1)
    {
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(path);
        long number = 0;
        if (match.Success)
        {
            number = long.Parse(match.Value);
        }
        return number;
    }
    public static void connectedComponents(List<long> vertice)
    {
        //avgsim
        double frac3 = 0;
        double[] avg = new double[vertice.Count];
        //no of component
        long x = 0;
        // components
        l6 = new List<long>[vertice.Count];
        Queue<long> q = new Queue<long>();
        //no of rows
        float c = 0;
        foreach (var o in vertice)  //o (V+E)
        {
            if (visited2[o] == true)
            {
                continue;
            }
            else
            {
                l6[x] = new List<long>();
                l6[x].Add(o);
                q.Enqueue(o);
                visited2[o] = true;
            }
            while (q.Count != 0)
            {
                foreach (var neighbor in d[q.Dequeue()])
                {
                    avg[x] += neighbor.wight;
                    c++;
                    if (visited2[neighbor.vertices] == true)
                        continue;
                    else
                    {
                        l6[x].Add(neighbor.vertices);
                        visited2[neighbor.vertices] = true;
                        q.Enqueue(neighbor.vertices);
                    }
                }
            }
            if (c != 0)
                avg[x] /= c;
            //to sort according to maxw
            if(l5.Contains(avg[x] + frac3))
                frac3 += 0.0000000000001;
            indexOfConnComp[avg[x] + frac3] = x;
            l5.Add(avg[x] + frac3);
            x++;
            c = 0;
        }
        string filename2 = $"C:\\Users\\User\\Downloads\\Sample\\Medium\\Stat_file2.xlsx";
        WriteDataToExcelWithEpplus(filename2, x, l6, avg);
    }
    public static (List<KeyValuePair<string, string>>, List<KeyValuePair<string, string>>, List<double>) Mst(string filename)
    {
        list1.Sort();//e log e
        for (int i = list1.Count - 1; i >= 0; i--)//o(e+v)
        {
            foreach (var path in d2[list1[i]])
            {
                if (id[path.s1] == id[path.s2]) //takes vertex and gives id 
                {
                    list3.Add(path);
                }
                else
                {
                    foreach (var a in visited4[id[path.s2]])  //takes id gives list of ver2
                    {
                        id[a] = id[path.s1];
                        visited4[id[path.s1]].Add(a);
                    }
                }
            }
        }
        List<KeyValuePair<string, string>> col1 = new List<KeyValuePair<string, string>>();
        List<KeyValuePair<string, string>> col2 = new List<KeyValuePair<string, string>>();
        List<double> col3 = new List<double>();
        using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.First();

            double frac2 = 0;
            int p = 0;
            for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)//o(n) not sure
            {
                int count = l5.Count;
                string l = (string)worksheet.Cells[row, 1].Value;
                string l2 = (string)worksheet.Cells[row, 2].Value;
                if (worksheet.Cells[row, 3].Value == null)
                    continue;
                double l3 = (double)worksheet.Cells[row, 3].Value;
                string hyper1 = "";
                string hyper2 = "";
                if (worksheet.Cells[row, 1].Hyperlink != null && worksheet.Cells[row, 2].Hyperlink != null)
                {
                    p++;
                    hyper1 = worksheet.Cells[row, 1].Hyperlink.AbsoluteUri;
                    hyper2 = worksheet.Cells[row, 2].Hyperlink.AbsoluteUri;
                }
                else
                {
                    hyper1 = (string)worksheet.Cells[row, 1].Value;
                    hyper2 = (string)worksheet.Cells[row, 2].Value;
                }
                long extractedpath2 = ext(l2);
                long extractedpath = ext(l);
                ver2 v = new ver2();
                v.s1 = extractedpath;
                v.s2 = extractedpath2;
                double p1 = extper(l);
                double p2 = extper(l2);
                if (list3.Contains(v)) //rejected
                {
                    continue;
                }
                foreach (var it in l5)
                {
                    if (l6[indexOfConnComp[it]].Contains(v.s1))
                    {
                        if (list2.Contains(count + frac2 + l3 / 5000 + Math.Max(p1, p2) / 50000000))
                            frac2 -= 0.0000000000001;
                        d3[count + frac2 + l3 / 5000 + Math.Max(p1, p2) / 50000000] = new List<ver3>();
                        v3.f1 = l;
                        v3.f2 = l2;
                        v3.hyperlink1 = hyper1;
                        v3.hyperlink2 = hyper2;
                        v3.lineMatch = l3;
                        d3[count + frac2 + l3 / 5000 + Math.Max(p1, p2) / 50000000].Add(v3);
                        list2.Add(count + frac2 + l3 / 5000 + Math.Max(p1, p2) / 50000000);
                    }
                    count--;
                }
            }
            Console.WriteLine(p);
            list2.Sort();//o(elog e)
            for (int i = list2.Count - 1; i >= 0; i--)//o(e) without edges of cycle
            {
                foreach (var rowImformation in d3[list2[i]])
                {
                    col1.Add(new KeyValuePair<string, string>(rowImformation.f1, rowImformation.hyperlink1));
                    col2.Add(new KeyValuePair<string, string>(rowImformation.f2, rowImformation.hyperlink2));
                    col3.Add(rowImformation.lineMatch);
                }
            }
        }
        return (col1, col2, col3);
    }
    public static void write_to_excel(string filename, List<KeyValuePair<string, string>> col1, List<KeyValuePair<string, string>> col2, List<double> col3, string f, string f2, string f3)
    {
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            worksheet.Cells[1, 1].Value = f;
            worksheet.Cells[1, 2].Value = f2;
            worksheet.Cells[1, 3].Value = f3;
            for (int i = 0; i < col1.Count; i++)//o(n)
            {
                //    string escapedUrl = Uri.EscapeUriString(col1[i]);
                //    string escapedUrl2 = Uri.EscapeUriString(col2[i]);
                if (col1[i].Value.Length != 0 && col2[i].Value.Length != 0)
                {
                    worksheet.Cells[i + 2, 1].Hyperlink = new Uri(col1[i].Value);
                    worksheet.Cells[i + 2, 2].Hyperlink = new Uri(col2[i].Value);
                }

                worksheet.Cells[i + 2, 1].Value = col1[i].Key;
                worksheet.Cells[i + 2, 2].Value = col2[i].Key;
                worksheet.Cells[i + 2, 3].Value = col3[i];
                var cellStyle = worksheet.Cells[i + 2, 1].Style;
                cellStyle.Font.Color.SetColor(System.Drawing.Color.Blue);
                cellStyle.Font.UnderLine = true;
                var cellStyle2 = worksheet.Cells[i + 2, 2].Style;
                cellStyle2.Font.Color.SetColor(System.Drawing.Color.Blue);
                cellStyle2.Font.UnderLine = true;
            }

            worksheet.Cells.AutoFitColumns();


            package.SaveAs(filename);
        }
    }
    public static void Main(string[] args)
    {
        var watch3 = Stopwatch.StartNew();
        string filename = "C:\\Users\\User\\Downloads\\Sample\\Medium\\2-Input.xlsx";
        (List<string> col1Data, List<string> col2Data, List<double> col3Data) = ReadDataToLists(filename);
        var watch = Stopwatch.StartNew();
        connectedComponents(vertice);
        watch.Stop();
        Console.WriteLine(
              $"The Execution time of the stat is {watch.ElapsedMilliseconds}ms");
        var watch2 = Stopwatch.StartNew();

        (List<KeyValuePair<string, string>> colData1, List<KeyValuePair<string, string>> colData2, List<double> colData3) = Mst(filename);
        write_to_excel($"C:\\Users\\User\\Downloads\\Sample\\Medium\\MST_File2.xlsx", colData1, colData2, colData3, "File 1", "File 2", "Lines Matched");
        watch2.Stop();
        watch3.Stop();
        Console.WriteLine(
              $"The Execution time of the Mst is {watch2.ElapsedMilliseconds}ms");
        Console.WriteLine(
             $"The Execution time of the total is {watch3.ElapsedMilliseconds}ms");
    }
}

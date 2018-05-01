using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace miningReporter
{
    using iTextSharp.text;
    using Pdfizer;
    using System.Data;
    using System.Data.SqlClient;

    class Program
    {
        public static SqlConnection baglanti;


        public static DataTable SqlDT(string query)
        {




            SqlCommand cmd = new SqlCommand(query, baglanti);



            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
            SqlDataReader dr = cmd.ExecuteReader();

            var tb = new DataTable();
            tb.Load(dr);

            baglanti.Close();


            return tb;
        }

        public static void insert(string query)
        {
            SqlCommand cmd = new SqlCommand(query, baglanti);
            /* SqlCommand cmd = new SqlCommand("INSERT INTO  (dt,hid,home,aid,aa,f1,fx,ff,zz,xx) VALUES (getdate(),@a,@b,@c,@d,@o1,@o2,@o3,@o4,@o5)", baglanti);
          cmd.Parameters.AddWithValue("@a", a);
            cmd.Parameters.AddWithValue("@b", b);
            cmd.Parameters.AddWithValue("@c", c);
            cmd.Parameters.AddWithValue("@d", d);
            cmd.Parameters.AddWithValue("@o1", o1);
            cmd.Parameters.AddWithValue("@o2", o2);
            cmd.Parameters.AddWithValue("@o3", o3);
            cmd.Parameters.AddWithValue("@o4", o4);
            cmd.Parameters.AddWithValue("@o5", o5);*/
            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }

            cmd.ExecuteNonQuery();
            baglanti.Close();
        }
        static void Main(string[] args)
        {
            baglanti = new SqlConnection("Data Source=.; Initial Catalog=chatbot; Integrated Security=true");
            string query = "SELECT * FROM chatbot.dbo.MyTable_rpr_allWords";
            DataTable dt = SqlDT(query);
            string sb = HTMLTableString(dt);
            to_pdf(sb);
        }

        static void to_pdf(string sb)
        {
            String success = "SUCCESS";
            HtmlToPdfConverter html2pdf = new HtmlToPdfConverter();

            try
            {
                String outputFilename = "output.pdf";
                File.Delete(outputFilename);
                FileStream outputWriter = new FileStream(outputFilename, FileMode.OpenOrCreate);
                html2pdf.Open(outputWriter);
                html2pdf.AddChapter(@"Dummy Chapter");
                //    html2pdf.Run(sb.ToString());
                html2pdf.Run(sb.ToString());
                html2pdf.AddChapter(@"A Wiki page");
                html2pdf.ImageUrl = true;

                // TextReader reader = new StreamReader("input.htm");

                //html2pdf.Run(reader.ReadToEnd());

                //html2pdf.AddChapter(@"Boost page");
                html2pdf.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                success = "FAIL";
            }
            Console.WriteLine(success);
            Console.ReadLine();
        }

        public static string HTMLTableString(DataTable dt)
        {
            // string strM = "EXEC KKBSITE_SP_FINANSAL " + ViewState["base"].ToString();
            String RVl = "";
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("<table border=\"0\" cellpadding=\"0\" cellspacing=\"1\" ><thead><tr>");
                foreach (DataColumn c in dt.Columns)
                {
                    sb.AppendFormat("<th>{0}</th>", c.ColumnName);
                }
                sb.AppendLine("</tr></thead><tbody>");

                foreach (DataRow dr in dt.Rows)
                {
                    sb.Append("<tr>"); foreach (object o in dr.ItemArray)
                    {
                        sb.AppendFormat("<td>{0}</td>", o.ToString());
                        //System.Web.HttpUtility.HtmlEncode());
                    } sb.AppendLine("</tr>");
                } sb.AppendLine("</tbody></table>");
                RVl = sb.ToString();
            }
            catch (Exception ex)
            {
                RVl = "HATA @ConvertDataTable2HTMLString: " + ex;//  Page.ClientScript.RegisterStartupScript(typeof(Page), "bisey3", "alert('bunu Alper Özen e gönderiniz\n" + strM + "\n" + ex.ToString() + "');", true);
            }
            return RVl;
        }
        public static DataTable HTMLTransposedTable(DataTable inputTable)
        {
            DataTable outputTable = new DataTable();
            // Add columns by looping rows
            // Header row's first column is same as in inputTable
            outputTable.Columns.Add(inputTable.Columns[0].ColumnName.ToString());
            // Header row's second column onwards, 'inputTable's first column taken
            foreach (DataRow inRow in inputTable.Rows)
            {
                string newColName = inRow[0].ToString();
                outputTable.Columns.Add(newColName);
            }
            // Add rows by looping columns       
            for (int rCount = 1; rCount <= inputTable.Columns.Count - 1; rCount++)
            {
                DataRow newRow = outputTable.NewRow();
                // First column is inputTable's Header row's second column
                newRow[0] = inputTable.Columns[rCount].ColumnName.ToString();
                for (int cCount = 0; cCount <= inputTable.Rows.Count - 1; cCount++)
                {
                    string colValue = inputTable.Rows[cCount][rCount].ToString();
                    newRow[cCount + 1] = colValue;
                }
                outputTable.Rows.Add(newRow);
            }
            return outputTable;

        }

    }
}

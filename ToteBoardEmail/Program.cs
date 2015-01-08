using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Hps.Core;
using System.Data.SqlClient;

namespace ToteBoardEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            string to = DBConfigManager.AppSettings["ToteboardEmailAddresses"].ToString();
            //string to = "Zac.Gragg@e-hps.com,Kenton.Cissell@e-hps.com,Chad.Hinton@e-hps.com,Tony.Capucille@e-hps.com";
            //string to = "joseph.feist@e-hps.com";
            string from = "ToteBoardEmail@e-hps.com";
            string subject = "Daily Toteboard Email";

            string emailBody = string.Empty;

            if (DateTime.Now.Day <= 10)
            {
                DateTime now = DateTime.Now;
                DateTime then = now.AddMonths(-1);

                DateTime beginDate1 = new DateTime(then.Year, then.Month, 1);
                DateTime endDate1 = new DateTime(then.Year, then.Month, DateTime.DaysInMonth(then.Year, then.Month));

                DateTime beginDate2 = new DateTime(now.Year, now.Month, 1);
                DateTime endDate2 = new DateTime(now.Year, now.Month, DateTime.DaysInMonth(now.Year, now.Month));

                emailBody = "<h2>Report for " + System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(then.Month) + " " + then.Year.ToString() + "</h2>" + GetEmailTableHTML(beginDate1, endDate1) + "<h2>Report for " + System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(now.Month) + " " + now.Year.ToString() + "</h2>" + GetEmailTableHTML(beginDate2, endDate2);
            }
            else
            {
                DateTime now = DateTime.Now;

                DateTime beginDate = new DateTime(now.Year, now.Month, 1);
                DateTime endDate = new DateTime(now.Year, now.Month, DateTime.DaysInMonth(now.Year, now.Month));

                emailBody = "<h2>Report for " + System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(now.Month) + " " + now.Year.ToString() + "</h2>" + GetEmailTableHTML(beginDate, endDate);
            }

            using (System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage(from, to, subject, emailBody))
            {
                mail.IsBodyHtml = true;

                try
                {
                    System.Net.Mail.SmtpClient clnt;
                    clnt = new System.Net.Mail.SmtpClient(DBConfigManager.AppSettings["SmtpServer"].ToString());
                    clnt.Send(mail);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Email was NOT succesfully sent! Email Send Failed with exception:" + Environment.NewLine + e.Message.Trim());
                }
            }
        }

        public static string GetEmailTableHTML(DateTime beginDate, DateTime endDate)
        {
            string regions = GetRegions();
            double autoRevTotal = 0;
            double teamTotal = 0;
            double goalTotal = 0;
            string goalPercTotal = string.Empty;

            List<SqlParameter> param = new List<SqlParameter>();
            param.Add(new SqlParameter("@MonthEnd", endDate));

            string table = "<table border=1><tr><td><b>Region</b></td><td><b>AutoReview Margin</b></td><td><b>Team Margin</b></td><td><b>Goal Amount</b></td><td><b>% Goal Achieved</b></td></tr>";

            using (DataTable dt = DAP.Connections.Get("DMSales").ExecuteDataTable(CommandType.Text, @"--Margin
                DECLARE @MonthBegin DATETIME

                SELECT @MonthBegin = DATEADD(month, DATEDIFF(month, 0, @MonthEnd), 0)

                DECLARE @BeginOfMonth INT
                DECLARE @EndOfMonth INT

                DECLARE @Month varchar(2)
                DECLARE @Day varchar(2)

                IF DATEPART (MM , @MonthEnd)<=9
	                BEGIN
		                SET @Month='0'+ CAST(DATEPART (MM , @MonthEnd) as char(1))
	                END 
                ELSE
	                BEGIN
		                SET @Month=CAST(DATEPART (MM , @MonthEnd)as char(2))
	                END
	
                IF DATEPART (DD , @MonthEnd)<=9
	                BEGIN
		                SET @Day='0'+ CAST(DATEPART (DD , @MonthEnd) as char(1))
	                END 
                ELSE
	                BEGIN
		                SET @Day=CAST(DATEPART (D , @MonthEnd)as CHAR(2))
	                END

                SELECT @EndOfMonth = CAST(DATEPART (YYYY , @MonthEnd) as char(4) ) +@Month +@Day

                IF DATEPART (MM , @MonthBegin)<=9
	                BEGIN
		                SET @Month='0'+ CAST(DATEPART (MM , @MonthBegin) as char(1))
	                END 
                ELSE
	                BEGIN
		                SET @Month=CAST(DATEPART (MM , @MonthBegin)as char(2))
	                END
	
                IF DATEPART (DD , @MonthBegin)<=9
	                BEGIN
		                SET @Day='0'+ CAST(DATEPART (DD , @MonthBegin) as char(1))
	                END 
                ELSE
	                BEGIN
		                SET @Day=CAST(DATEPART (D , @MonthBegin)as CHAR(2))
	                END

                SELECT @BeginOfMonth = CAST(DATEPART (YYYY , @MonthBegin) as char(4) ) +@Month +@Day

                SELECT CASE WHEN ex.ExclusionID IS NULL THEN Region
		                ELSE 'Other' END AS Region
	                , Final_Margin
                INTO #NetMargin
                FROM Fact.Sales (NOLOCK) fs
                JOIN Dim.Employee (NOLOCK) e ON fs.EmployeeKey = e.EmployeeKey
                JOIN dim.Geography (NOLOCK) g on e.GeographyKey = g.GeographyKey
                LEFT JOIN Exclusions (NOLOCK) ex ON e.ID = ex.EmployeeID AND GETDATE() BETWEEN ex.StartDate AND ex.EndDate AND ex.Class = 'Other'
                WHERE fs.Final_StartDate BETWEEN @BeginOfMonth AND @EndOfMonth

                SELECT Region
	                , SUM(Final_Margin) Final_Margin
                INTO #NetMarginSum
                FROM #NetMargin
                GROUP BY Region

                SELECT CASE WHEN ex.ExclusionID IS NULL THEN Region
		                ELSE 'Other' END AS Region
	                , AutoRev_Margin
                INTO #AutoRev
                FROM Fact.Sales (NOLOCK) fs
                JOIN Dim.Employee (NOLOCK) e ON fs.EmployeeKey = e.EmployeeKey
                JOIN dim.Geography (NOLOCK) g on e.GeographyKey = g.GeographyKey
                LEFT JOIN Exclusions (NOLOCK) ex ON e.ID = ex.EmployeeID AND GETDATE() BETWEEN ex.StartDate AND ex.EndDate AND ex.Class = 'Other'
                WHERE fs.AutoRev_StartDate BETWEEN @BeginOfMonth AND @EndOfMonth
                AND fs.AutoRev_EndDate IS NULL

                SELECT Region
	                , SUM(AutoRev_Margin) AutoRev_Margin
                INTO #AutoRevSum
                FROM #AutoRev
                GROUP BY Region

                --Sales Goals
                SELECT g.Region
	                , SUM(sg.Margin) SalesGoal
                INTO #SalesGoals
                FROM Fact.SalesGoals (NOLOCK) sg
                JOIN dim.Geography (NOLOCK) g ON sg.GeographyKey = g.GeographyKey
                WHERE g.IsActive = 1
                AND @EndOfMonth BETWEEN sg.StartDateKey AND EndDateKey
                GROUP BY g.Region

                --Final Select
                SELECT nm.Region
	                , nm.Final_Margin
	                , ar.AutoRev_Margin
	                , sg.SalesGoal
                FROM #NetMarginSum nm
                LEFT JOIN #AutoRevSum ar ON nm.Region = ar.Region
                LEFT JOIN #SalesGoals sg ON nm.Region = sg.Region
                ORDER BY nm.Final_Margin DESC

                DROP TABLE #NetMargin
                DROP TABLE #AutoRev
                DROP TABLE #NetMarginSum
                DROP TABLE #AutoRevSum
                DROP TABLE #SalesGoals", param.ToArray()))
            {
                foreach (DataRow dr in dt.Rows)
                {
                    double autoD = Convert.ToDouble(dr["AutoRev_Margin"].ToString() == "" ? "0" : dr["AutoRev_Margin"].ToString());
                    string autoS = autoD.ToString("C");
                    double teamD = Convert.ToDouble(dr["Final_Margin"].ToString() == "" ? "0" : dr["Final_Margin"].ToString());
                    string teamS = teamD.ToString("C");
                    double goalD = Convert.ToDouble(dr["SalesGoal"].ToString() == "" ? "0" : dr["SalesGoal"].ToString());
                    
                    //Goals do not exist for Other and HSC so they are settings for this app only
                    if (dr["Region"].ToString() == "Other")
                    {
                        goalD = Convert.ToDouble(DBConfigManager.AppSettings["GoalOther"].ToString().Replace(",", ""));
                    }
                    if (dr["Region"].ToString() == "HSC")
                    {
                        goalD = Convert.ToDouble(DBConfigManager.AppSettings["GoalHSC"].ToString().Replace(",", ""));
                    }
                    
                    string goalS = goalD.ToString("C");
                    string goalPerc = goalD == 0.00 ? string.Empty : (teamD / goalD).ToString("P2");

                    autoRevTotal += autoD;
                    teamTotal += teamD;
                    goalTotal += goalD;

                    table += "<tr><td>" + dr["Region"].ToString() + "</td><td>" + autoS + "</td><td>" + teamS + "</td><td>" + goalS + "</td><td>" + goalPerc + "</td></tr>";
                }
            }

            goalPercTotal = (teamTotal / goalTotal).ToString("P2");

            table += "<tr><td><b>Total</b></td><td><b>" + autoRevTotal.ToString("C") + "</b></td><td><b>" + teamTotal.ToString("C") + "</b></td><td><b>" + goalTotal.ToString("C") + "</b></td><td><b>" + goalPercTotal + "</b></td></tr></table>";

            return table;
        }

        public static string GetRegions()
        {
            DataTable dt = DAP.Connections.Get("DMSales").ExecuteDataTable(CommandType.Text, @"select distinct Region from dim.vw_Geography where IsActive = 1");

            string regions = string.Empty;

            foreach (DataRow dr in dt.Rows)
            {
                string region = dr["Region"].ToString();

                if (region == "Unknown")
                {
                    region = "Other";
                }

                regions += region + ",";
            }

            return regions;
        }
    }
}

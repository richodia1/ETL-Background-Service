//Initial code author is Richard Okosodo
//On his interview stage with Arik Air Line,
//A place where everybody want to belong.
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArikInterview
{
   public class DataAccessLayer
    {
        public bool insert1010RecordType(String line)
        {
            string PNRnumber = line.Substring(5, 6);
            string TSMdocumenttype = line.Substring(11, 1);
            string TSMsubtypeManual = line.Substring(12, 1);
            string Ticketissuingairline = line.Substring(13, 1);
            string PTAcreationdate = line.Substring(15, 7);
            string PTAcreationlocation = line.Substring(22, 3);
            string PTAcreationagentID = line.Substring(25, 2);
            string SalesCreditorIATA = line.Substring(27, 8);
            string EMDindicator = line.Substring(35, 1);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PNR number],[TSM document type - P/M/E IND],[TSM subtype - Manual /Auto Issued],[Ticket issuing airline],[PTA creation date],[PTA creation location],[PTA creation agent ID],[Sales Creditor - IATA],[EMD indicator]) " +
                     "VALUES(@value1, @value2, @value3, @value4,@value5,@value6, @value7, @value8, @value9)", cn);
                cmd1.Parameters.AddWithValue("@value1", PNRnumber);
                cmd1.Parameters.AddWithValue("@value2", TSMdocumenttype);
                cmd1.Parameters.AddWithValue("@value3", TSMsubtypeManual);
                cmd1.Parameters.AddWithValue("@value4", Ticketissuingairline);
                cmd1.Parameters.AddWithValue("@value5", PTAcreationdate);
                cmd1.Parameters.AddWithValue("@value6", PTAcreationlocation);
                cmd1.Parameters.AddWithValue("@value7", PTAcreationagentID);
                cmd1.Parameters.AddWithValue("@value8", SalesCreditorIATA);
                cmd1.Parameters.AddWithValue("@value9", EMDindicator);
                cmd1.ExecuteNonQuery();
            }
            return true;

        }
        public bool insert1015RecordType(String line)
        {
            string TKT = line.Substring(5, 5);
            string IssueAircode = line.Substring(10, 2);
            string IssueCitycode = line.Substring(12, 3);
            string IssueAgentcode = line.Substring(15, 2);
            string TSMSalesIndicator = line.Substring(17, 1);
            string IssueLocaldate = line.Substring(18, 7);
            string IssueSystemdate = line.Substring(25, 7);
            string IssueLocalTiming = line.Substring(32, 8);
            string IssueSystemTime = line.Substring(40, 8);
            string ReportingType = line.Substring(48, 1);
            string ReportingIndicator = line.Substring(49, 1);
            string ReportingCity = line.Substring(50, 3);
            string ReportingAgentID = line.Substring(53, 2);
            string CSRNumber = line.Substring(55, 6);
            string TicketingMode = line.Substring(61, 1);
            string DataReciever = line.Substring(62, 2);
            string IssueAirLinecode = line.Substring(64, 2);
            string ServicingAirlineCode = line.Substring(66, 4);
            string AirLineNumericCodeCheckDigit = line.Substring(71, 1);
            string AirLineNumericCode = line.Substring(72, 3);
            string FormCode = line.Substring(75, 3);
            string SerialNumberFrom = line.Substring(78, 7);
            string SerialNumberTo = line.Substring(85, 7);
            string CheckDigits = line.Substring(92, 1);
            string CheckDigitsOfExchangeCoupon = line.Substring(93, 1);
            string NumberGigitIssueMDA = line.Substring(94, 2);
            string StockControlNumber = line.Substring(96, 14);
            string Spares = line.Substring(110, 10);
            string FullName = line.Substring(120);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([TKT issuing ATB printer symbolic Addr],[Issuing airlinecode],[Issuing city code],[Issuing agent code],[TSM sales indicator],[Issuing local datetime],[Issuingsystemdatetime],[IssuingsystemLocaltime],[Issuing system time],[Reporting type],[Reported indicator],[Reporting city],[Reporting agent id],[CSR number],[Ticketing mode],[Data receiver],[Issuing Airline code],[Servicing airline code],[Airline numeric code check digit],[Airline numeric code],[Form code],[Serial number 'from'],[Serial number 'to'],[Checkdigit],[Check digit of exchange coupon],[Number of coupons issued this MDA],[Stock control number including check digit],[Spares],[Full Airline name]) " +
                     //,,[Issuing Airline code],[Servicing airline code],[Airline numeric code check digit],[Airline numeric code],[Form code],[Serial number 'from'],[Serial number 'to'],[Checkdigit],[Check digit of exchange coupon],[Number of coupons issued this MDA],[Stock control number including check digit],[Spares],[Full Airline name]) " +
                     "VALUES(@value1, @value2, @value3, @value4,@value5,@value6,@value7,@value8,@value9,@value10,@value11,@value12,@value13,@value14,@value15,@value16,@value17,@value18,@value19,@value20,@value21,@value22,@value23,@value24,@value25,@value26,@value27,@value28,@value29)", cn);
                //,,,, @value7, @value8, @value9, @value10, @value11, @value12, @value13, @value14, @value15, @value16,@value17, @value18, @value19, @value20, @value21, @value22, @value23, @value24, @value25, @value26, @value27, @value28, @value29)", cn);
                cmd1.Parameters.AddWithValue("@value1", TKT);
                cmd1.Parameters.AddWithValue("@value2", IssueAircode);
                cmd1.Parameters.AddWithValue("@value3", IssueCitycode);
                cmd1.Parameters.AddWithValue("@value4", IssueAgentcode);
                cmd1.Parameters.AddWithValue("@value5", TSMSalesIndicator);
                cmd1.Parameters.AddWithValue("@value6", IssueLocaldate);
                cmd1.Parameters.AddWithValue("@value7", IssueSystemdate);
                cmd1.Parameters.AddWithValue("@value8", IssueLocalTiming);
                cmd1.Parameters.AddWithValue("@value9", IssueSystemTime);
                cmd1.Parameters.AddWithValue("@value10", ReportingType);
                cmd1.Parameters.AddWithValue("@value11", ReportingIndicator);
                cmd1.Parameters.AddWithValue("@value12", ReportingCity);
                cmd1.Parameters.AddWithValue("@value13", ReportingAgentID);
                cmd1.Parameters.AddWithValue("@value14", CSRNumber);
                cmd1.Parameters.AddWithValue("@value15", TicketingMode);
                cmd1.Parameters.AddWithValue("@value16", DataReciever);
                cmd1.Parameters.AddWithValue("@value17", IssueAirLinecode);
                cmd1.Parameters.AddWithValue("@value18", ServicingAirlineCode);
                cmd1.Parameters.AddWithValue("@value19", AirLineNumericCodeCheckDigit);
                cmd1.Parameters.AddWithValue("@value20", AirLineNumericCode);
                cmd1.Parameters.AddWithValue("@value21", FormCode);
                cmd1.Parameters.AddWithValue("@value22", SerialNumberFrom);
                cmd1.Parameters.AddWithValue("@value23", SerialNumberTo);
                cmd1.Parameters.AddWithValue("@value24", CheckDigits);
                cmd1.Parameters.AddWithValue("@value25", CheckDigitsOfExchangeCoupon);
                cmd1.Parameters.AddWithValue("@value26", NumberGigitIssueMDA);
                cmd1.Parameters.AddWithValue("@value27", StockControlNumber);
                cmd1.Parameters.AddWithValue("@value28", Spares);
                cmd1.Parameters.AddWithValue("@value29", FullName);
                cmd1.ExecuteNonQuery();
               
            }
            return true;
        }
        public bool insert101FRecordType(String line)
        {
            string PassNameItem = line.Substring(5, 2);
            string PassNameInitial = line.Substring(7, 6);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Passenger Name Item],[Passenger Name/Initial]) " +
                     "VALUES(@value1, @value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", PassNameItem);
                cmd1.Parameters.AddWithValue("@value2", PassNameInitial);
                cmd1.ExecuteNonQuery();
                
            }
            return true;
        }
        public bool insert1025RecordType(String line)
        {
            string edMdata = line.Substring(5, 10);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Endorsement data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3150RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Pax Description String]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3200RecordType(String line)
        {
            string edMdata = line.Substring(5,1);
            string edMdata1 = line.Substring(6);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Passenger Name PRINT status],[Passenger Name PRINT status]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.Parameters.AddWithValue("@value1", edMdata1);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3700RecordType(String line)
        {
            string edMdata = line.Substring(5, 5);
            string edMdata1 = line.Substring(10);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA ticket tax status],[Ticket tax currency amount]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.Parameters.AddWithValue("@value1", edMdata1);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3750RecordType(String line)
        {
            string edMdata = line.Substring(5, 3);
            string edMdata1 = line.Substring(8);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Ticket PFC airport code/amount status],[PFC airport code/amount data]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.Parameters.AddWithValue("@value1", edMdata1);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3225RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA reservation data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3230RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA Endorsement data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3315RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Passenger Address data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3330RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA sponsor data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3400RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA Ticket total data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3420RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([TKT fare construction]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3500RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([TKT fare/tax data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert3600RecordType(String line)
        {
            string edMdata = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA original form of payment]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", edMdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }

        //This is 1500 data type handler
        public bool insert1500RecordType(String line)
        {
            string PFPCode = line.Substring(5, 5);
            string FOPSN = line.Substring(10, 2);
            string FOPSC = line.Substring(21, 10);
            string FOPSC2 = line.Substring(31, 5);
            string FOPPP = line.Substring(36, 0);
           
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Primary form of Payment Code],[FOP subitem number],[FOP subitem code1],[FOP subitem code],[FOP Print primary]) " +
                     "VALUES(@value1, @value2, @value3, @value4, @value5)", cn);
                cmd1.Parameters.AddWithValue("@value1", PFPCode);
                cmd1.Parameters.AddWithValue("@value2", FOPSN);
                cmd1.Parameters.AddWithValue("@value3", FOPSC);
                cmd1.Parameters.AddWithValue("@value4", FOPSC2);
                cmd1.Parameters.AddWithValue("@value5", FOPPP);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1502 data type handler
        public bool insert1502RecordType(String line)
        {
            string PFP = line.Substring(5, 14);
            string FOPSN = line.Substring(19, 2);
            string FOPSC1 = line.Substring(21, 6);
            string FOPSC2 = line.Substring(27, 6);
            string FOPSD = line.Substring(33, 3);

            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Primary Form of payment],[FOP Subitem number],[FOP Subitem code1],[FOP Subitem code],[FOP Subitem Data]) " +
                     "VALUES(@value1,@value2, @value3, @value4, @value5)", cn);
                cmd1.Parameters.AddWithValue("@value1", PFP);
                cmd1.Parameters.AddWithValue("@value2", FOPSN);
                cmd1.Parameters.AddWithValue("@value3", FOPSC1);
                cmd1.Parameters.AddWithValue("@value4", FOPSC2);
                cmd1.Parameters.AddWithValue("@value5", FOPSD);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1510 data type handler

        public bool insert1510RecordType(String line)
        {
            string TKTotal = line.Substring(5, 5);
            
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([TKT total]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", TKTotal);
               
                cmd1.ExecuteNonQuery();

            }
            return true;
        }
        //This is 1400 data type handler
        public bool insert1400RecordType(String line)
        {
            string Cvalue = line.Substring(5, 5);
            string Camount = line.Substring(10, 2);
            string Cretained = line.Substring(21, 10);
            string Crating = line.Substring(31, 5);
           

            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Commission value],[Commissionable amount],[Commission retained],[Commission rate]) " +
                     "VALUES(@value1,@value2, @value3, @value4, @value5)", cn);
                cmd1.Parameters.AddWithValue("@value1", Cvalue);
                cmd1.Parameters.AddWithValue("@value2", Camount);
                cmd1.Parameters.AddWithValue("@value3", Cretained);
                cmd1.Parameters.AddWithValue("@value4", Crating);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert5005RecordType(String line)
        {
            string Cvalue = line.Substring(5, 2);
            string Camount = line.Substring(7, 3);
            string Cretained = line.Substring(10, 8);
            string Crating = line.Substring(18);


            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA issuing airline],[PTA issuing city code],[Issuing IATA code],[Document data of manually issued PTA doc.]) " +
                     "VALUES(@value1,@value2, @value3, @value4, @value5)", cn);
                cmd1.Parameters.AddWithValue("@value1", Cvalue);
                cmd1.Parameters.AddWithValue("@value2", Camount);
                cmd1.Parameters.AddWithValue("@value3", Cretained);
                cmd1.Parameters.AddWithValue("@value4", Crating);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1512 data type handler
        public bool insert1512RecordType(String line)
        {
            string NetF = line.Substring(5, 4);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([NETT fare]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", NetF);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1700 data type handler
        public bool insert1700RecordType(String line)
        {
            string TCDS = line.Substring(5, 1);
            string TCD = line.Substring(6, 1);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Tour code data source],[Tour code data]) " +
                     "VALUES(@value1, @value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", TCDS);
                cmd1.Parameters.AddWithValue("@value2", TCD);
                cmd1.ExecuteNonQuery();

            }
            return true;
        }
        //This is 1800 data type handler
        public bool insert1800RecordType(String line)
        {
            string MDAtotal = line.Substring(5, 5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([MDA Auto Total]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", MDAtotal);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1802 data type handler
        public bool insert1802RecordType(String line)
        {
            string MDAMtotal = line.Substring(5, 4);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([MDA Manual Total(1802type)]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", MDAMtotal);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1804 data type handler
        public bool insert1804RecordType(String line)
        {
            string NetMT = line.Substring(5, 4);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Net Manual Total]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", NetMT);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1806 data type handler
        public bool insert1806RecordType(String line)
        {
            string Att = line.Substring(5, 4);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([NET Auto total]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", Att);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1850 data type handler
        public bool insert1850RecordType(String line)
        {
            string Tstatus = line.Substring(5, 2);
            string Tdata = line.Substring(5, 2);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Tax status],[Tax Data]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1",  Tstatus);
                cmd1.Parameters.AddWithValue("@value2", Tdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1900 data type handler
        public bool insert1900RecordType(String line)
        {
            string Estatus = line.Substring(5, 2);
            string Edata = line.Substring(5, 2);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Equivalent fare status)],[Equivalent fare data]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", Estatus);
                cmd1.Parameters.AddWithValue("@value2", Edata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert1902RecordType(String line)
        {
            string NetMT = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Net Equi fare data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", NetMT);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert5015RecordType(String line)
        {
            string NetMT = line.Substring(5);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA short routing data]) " +
                     "VALUES(@value1)", cn);
                cmd1.Parameters.AddWithValue("@value1", NetMT);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is B100 data type handler
        public bool insertB100RecordType(String line)
        {
            string c1 = line.Substring(5, 3);
            string c2 = line.Substring(8, 4);
            string c3 = line.Substring(13, 5);
            string c4 = line.Substring(18, 3);
            string c5 = line.Substring(21, 2);
            string c6 = line.Substring(23, 25);
            string c7 = line.Substring(48, 5);
            string c8 = line.Substring(53, 3);
            string c9 = line.Substring(56, 2);
            string c10 = line.Substring(58, 25);
            string c11= line.Substring(83, 2);
            string c12 = line.Substring(85, 2);
            string c13= line.Substring(87, 1);
            string c14 = line.Substring(88, 1);
            string c15 = line.Substring(89, 1);
            string c16 = line.Substring(90, 13);
            string c17 = line.Substring(103, 1);
            string c18 = line.Substring(104, 1);
            string c19 = line.Substring(105, 2);
            string c20 = line.Substring(107, 2);
            string c21 = line.Substring(109, 7);
            string c22 = line.Substring(116, 4);
            string c23 = line.Substring(120, 5);
            string c24 = line.Substring(125, 1);
            string c25 = line.Substring(126, 3);
            string c26 = line.Substring(129, 8);
            string c27 = line.Substring(137, 8);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Segment line number],[AAA city of last segmnt updated],[BPT Stopover N. Permitted indicator],[Boardpoint airport code],[Boardpoint multi airport indicator],[Extended boardpoint name],[Offpoint stopover not permitted indicator],[Offpoint airport code],[Offpoint multi airport indicator],[Extended offpoint name],[Number of coupons issued],[Number of flight coupons per document number],[Coupon number],[Coupon ordinal number], [Total number of ATB coupons],[Segment document number],[Document number check digit],[Segment dame indicator],[Global direction],[Airline code],[Departure date],[Line status],[Excess baggae number],[Excess baggage unit of measurement],[Currency code],[Excess baggage rate],[Excess baggage segment amount]) " +
                     "VALUES(@value1, @value2, @value3, @value4,@value5,@value6, @value7, @value8, @value9, @value10, @value11, @value12, @value13, @value14, @value15, @value16, @value17, @value18, @value19, @value20, @value21, @value22, @value23, @value24, @value25, @value26, @value27)", cn);
                cmd1.Parameters.AddWithValue("@value1", c1);
                cmd1.Parameters.AddWithValue("@value2", c2);
                cmd1.Parameters.AddWithValue("@value3", c3);
                cmd1.Parameters.AddWithValue("@value4", c4);
                cmd1.Parameters.AddWithValue("@value5", c5);
                cmd1.Parameters.AddWithValue("@value6", c6);
                cmd1.Parameters.AddWithValue("@value7", c7);
                cmd1.Parameters.AddWithValue("@value8", c8);
                cmd1.Parameters.AddWithValue("@value9", c9);
                cmd1.Parameters.AddWithValue("@value10", c10);
                cmd1.Parameters.AddWithValue("@value11", c11);
                cmd1.Parameters.AddWithValue("@value12", c12);
                cmd1.Parameters.AddWithValue("@value13", c13);
                cmd1.Parameters.AddWithValue("@value14", c14);
                cmd1.Parameters.AddWithValue("@value15", c15);
                cmd1.Parameters.AddWithValue("@value16", c16);
                cmd1.Parameters.AddWithValue("@value17", c17);
                cmd1.Parameters.AddWithValue("@value18", c18);
                cmd1.Parameters.AddWithValue("@value19", c19);
                cmd1.Parameters.AddWithValue("@value20", c20);
                cmd1.Parameters.AddWithValue("@value21", c21);
                cmd1.Parameters.AddWithValue("@value22", c22);
                cmd1.Parameters.AddWithValue("@value23", c23);
                cmd1.Parameters.AddWithValue("@value24", c24);
                cmd1.Parameters.AddWithValue("@value25", c25);
                cmd1.Parameters.AddWithValue("@value26", c26);
                cmd1.Parameters.AddWithValue("@value27", c27);
                
                cmd1.ExecuteNonQuery();

            }
            return true;
        }
        //This is B105 data type handler
        public bool insertB105RecordType(String line)
        {
            string Esb = line.Substring(5, 2);
            string res = line.Substring(5, 2);
            string fcta = line.Substring(5, 2);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([EXB fare calculation - Rate of Exchange)],[Rate of exchange source indicator],[Fare calculation Total amount]) " +
                     "VALUES(@value1,@value2,@value3)", cn);
                cmd1.Parameters.AddWithValue("@value1", Esb);
                cmd1.Parameters.AddWithValue("@value2", res);
                cmd1.Parameters.AddWithValue("@value3", res);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert5010RecordType(String line)
        {
            string Esb = line.Substring(5, 2);
            string res = line.Substring(7, 13);
            string fcta = line.Substring(20, 8);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Ticket issuing airline)],[Ticket document number from],[Ticket document number to]) " +
                     "VALUES(@value1,@value2,@value3)", cn);
                cmd1.Parameters.AddWithValue("@value1", Esb);
                cmd1.Parameters.AddWithValue("@value2", res);
                cmd1.Parameters.AddWithValue("@value3", res);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 1110 data type handler
        public bool insertB110RecordType(String line)
        {
            string d1 = line.Substring(5, 1);
            string d2 = line.Substring(5, 1);
            string d3 = line.Substring(7, 4);
            string d4 = line.Substring(12);
         

            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([EXB/MCO service type - Automatic/Manual Indicator],[Type of service IATA code/Input code],[Type of service group code],[Type of service ( The EXB/MCO is issued for )]) " +
                     "VALUES(@value1, @value2, @value3, @value4)", cn);
                cmd1.Parameters.AddWithValue("@value1", d1);
                cmd1.Parameters.AddWithValue("@value2", d2);
                cmd1.Parameters.AddWithValue("@value3", d3);
                cmd1.Parameters.AddWithValue("@value4", d4);
              
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        //This is 5030 data type handler
        public bool insert5030RecordType(String line)
        {
            string P1 = line.Substring(5, 7);
            string p2 = line.Substring(12, 3);
            string p3 = line.Substring(15, 2);
            string p4 = line.Substring(17, 2);
            string p5 = line.Substring(19, 4);
            string p6 = line.Substring(24);

            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Last update UTC date],[Last update city code],[Last update agent initial],[Last update airline code],[Last update UTC time],[MDA additional info text]) " +
                     "VALUES(@value1, @value2, @value3, @value4, @value5, @value6)", cn);
                cmd1.Parameters.AddWithValue("@value1", P1);
                cmd1.Parameters.AddWithValue("@value2", p2);
                cmd1.Parameters.AddWithValue("@value3", p3);
                cmd1.Parameters.AddWithValue("@value4", p4);
                cmd1.Parameters.AddWithValue("@value5", p5);
                cmd1.Parameters.AddWithValue("@value6", p6);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert501FRecordType(String line)
        {
            string P1 = line.Substring(5, 3);
            string p2 = line.Substring(8, 7);
            string p3 = line.Substring(15, 2);
            string p4 = line.Substring(17, 3);
            string p5 = line.Substring(20, 2);
            string p6 = line.Substring(22,4);

            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([PTA status code],[PTA status UTC],[Last update agent initial],[Last update city code],[Last update airline code],[UTC time (HHMM)]) " +
                     "VALUES(@value1, @value2, @value3, @value4, @value5, @value6)", cn);
                cmd1.Parameters.AddWithValue("@value1", P1);
                cmd1.Parameters.AddWithValue("@value2", p2);
                cmd1.Parameters.AddWithValue("@value3", p3);
                cmd1.Parameters.AddWithValue("@value4", p4);
                cmd1.Parameters.AddWithValue("@value5", p5);
                cmd1.Parameters.AddWithValue("@value6", p6);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert5020RecordType(String line)
        {
            string P1 = line.Substring(5, 3);
            string p2 = line.Substring(8, 2);
            string p3 = line.Substring(10, 2);
            string p4 = line.Substring(12, 3);
            string p5 = line.Substring(15, 2);
            string p6 = line.Substring(17, 2);

            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([City code],[Office function designator],[Airline alpha code],[City code],[Airline alpha code],[PARS queue number]) " +
                     "VALUES(@value1, @value2, @value3, @value4, @value5, @value6)", cn);
                cmd1.Parameters.AddWithValue("@value1", P1);
                cmd1.Parameters.AddWithValue("@value2", p2);
                cmd1.Parameters.AddWithValue("@value3", p3);
                cmd1.Parameters.AddWithValue("@value4", p4);
                cmd1.Parameters.AddWithValue("@value5", p5);
                cmd1.Parameters.AddWithValue("@value6", p6);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert5025RecordType(String line)
        {
            string P1 = line.Substring(5, 3);
            string p2 = line.Substring(8, 2);
            string p3 = line.Substring(10, 2);
            string p4 = line.Substring(12, 3);
            string p5 = line.Substring(15, 2);
            string p6 = line.Substring(17, 2);

            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([City code],[Office function designator],[Airline alpha code],[City code],[Airline alpha code],[PARS queue number]) " +
                     "VALUES(@value1, @value2, @value3, @value4, @value5, @value6)", cn);
                cmd1.Parameters.AddWithValue("@value1", P1);
                cmd1.Parameters.AddWithValue("@value2", p2);
                cmd1.Parameters.AddWithValue("@value3", p3);
                cmd1.Parameters.AddWithValue("@value4", p4);
                cmd1.Parameters.AddWithValue("@value5", p5);
                cmd1.Parameters.AddWithValue("@value6", p6);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insertD050RecordType(String line)
        {
            string Esb = line.Substring(5, 2);
            string res = line.Substring(7, 11);
            string fcta = line.Substring(18, 20);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Withhold tax rate)],[Withhold tax value],[Party Code]) " +
                     "VALUES(@value1,@value2,@value3)", cn);
                cmd1.Parameters.AddWithValue("@value1", Esb);
                cmd1.Parameters.AddWithValue("@value2", res);
                cmd1.Parameters.AddWithValue("@value3", res);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insertD000RecordType(String line)
        {
            string Tstatus = line.Substring(5, 13);
            string Tdata = line.Substring(18);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([MDA document number],[SCNR info]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", Tstatus);
                cmd1.Parameters.AddWithValue("@value2", Tdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insertB125RecordType(String line)
        {
            string Tstatus = line.Substring(5, 1);
            string Tdata = line.Substring(6);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([EXB fare calculation for PRINT - Automatic/Manual Indicator],[EXB fare calculation for PRINT data]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", Tstatus);
                cmd1.Parameters.AddWithValue("@value2", Tdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insertB120RecordType(String line)
        {
            string Tstatus = line.Substring(5, 1);
            string Tdata = line.Substring(6);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([EXB Routing for PRINT - Automatic/Manual Indicator],[EXB Routing for PRINT Data]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", Tstatus);
                cmd1.Parameters.AddWithValue("@value2", Tdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert7105RecordType(String line)
        {
            string Tstatus = line.Substring(5, 1);
            string Tdata = line.Substring(6);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([MCO service provider  - Automatic/Manual indicator],[Service provide name and location]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", Tstatus);
                cmd1.Parameters.AddWithValue("@value2", Tdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert7110RecordType(String line)
        {
            string Tstatus = line.Substring(5, 1);
            string Tdata = line.Substring(6);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([MDA printed],[Additional printed on MDA]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", Tstatus);
                cmd1.Parameters.AddWithValue("@value2", Tdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert7115RecordType(String line)
        {
            string Tstatus = line.Substring(5, 1);
            string Tdata = line.Substring(6);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([MDA number data],[MDA in connection with ticket number data]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", Tstatus);
                cmd1.Parameters.AddWithValue("@value2", Tdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
        public bool insert1852RecordType(String line)
        {
            string Tstatus = line.Substring(5, 3);
            string Tdata = line.Substring(8);
            string fileName = @"C:/test/Book2.xls";
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=0'", fileName);

            using (OleDbConnection cn = new OleDbConnection(connectionString))
            {
                cn.Open();
                OleDbCommand cmd1 = new OleDbCommand("INSERT INTO [Sheet1$] " +
                     "([Vat tax status],[Vat tax data]) " +
                     "VALUES(@value1,@value2)", cn);
                cmd1.Parameters.AddWithValue("@value1", Tstatus);
                cmd1.Parameters.AddWithValue("@value2", Tdata);
                cmd1.ExecuteNonQuery();
            }
            return true;
        }
    }
}











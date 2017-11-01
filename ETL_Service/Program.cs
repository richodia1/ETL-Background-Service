using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArikInterview
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.Title = "Arik Air ETL interview Task from RICHARD.";
            int counter = 0;
            string lines;

            // Read the file and display it line by line.
            System.IO.StreamReader file =
               new System.IO.StreamReader("C:/SecondDirectory/mpd261016.iss");
            while ((lines = file.ReadLine()) != null)
            {
                if (lines.StartsWith("1010") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1010RecordType(line);
                        Console.WriteLine(line);
                    }
                        
                   
                }
                if (lines.StartsWith("1015") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1015RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("101F") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert101FRecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1025") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1025RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1400") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1400RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1500") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1500RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1502") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1502RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1510") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1510RecordType(line);
                        Console.WriteLine(line);
                    }


                }
               
                if (lines.StartsWith("1512") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1512RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1700") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1700RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1800") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1800RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1802") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1802RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1804") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1804RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1806") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1806RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1850") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1850RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1852") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1852RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1900") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1900RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("1902") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert1902RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3150") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3150RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3200") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3200RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3225") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3225RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3230") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3230RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3315") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3315RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3330") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3330RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3400") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3400RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3420") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3420RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3500") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3500RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3600") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3600RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3700") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3700RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("3750") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert3750RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("5005") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert5005RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("5010") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert5010RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("5015") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert5015RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("501F") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert501FRecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("5020") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert5020RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("5025") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert5025RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("5030") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert5030RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("7105") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert7105RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("7110") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert7110RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("7115") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insert7115RecordType(line);
                        Console.WriteLine(line);
                    }
                }
                if (lines.StartsWith("B100") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insertB100RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("B105") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insertB105RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("B110") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insertB110RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("B120") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insertB120RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("B125") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insertD000RecordType(line);
                        Console.WriteLine(line);
                    }


                }
                if (lines.StartsWith("D000") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insertD000RecordType(line);
                        Console.WriteLine(line);
                    }


                }

                if (lines.StartsWith("D050") == true)
                {
                    ArrayList myAL = new ArrayList();
                    myAL.Add(lines);
                    foreach (string line in myAL)
                    {
                        DataAccessLayer db = new DataAccessLayer();
                        db.insertD050RecordType(line);
                        Console.WriteLine(line);
                    }


                }
               
                
               
                counter++;

            }

            file.Close();

            // Suspend the screen.
            Console.ReadLine();
        }
    }
}

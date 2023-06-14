using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using Oracle.ManagedDataAccess.Client;

namespace ClassLibrary1
{
    partial class DataSet1
    {


        private string connString;
        public string ConnString
        {
            get { return connString; }
            set { connString = value; }
        }

        OracleConnection myOraConnection;
        public OracleConnection MyOraConnection
        {
            get
            {
                if (this.myOraConnection == null)
                {
                    this.myOraConnection = new OracleConnection(ConnString);
                }
                return myOraConnection;
            }
            set { myOraConnection = value; }
        }

        static string ByteArrayToString(byte[] arrInput)
        {
            int i;
            StringBuilder sOutput = new StringBuilder(arrInput.Length);
            for (i = 0; i < arrInput.Length; i++)
            {
                sOutput.Append(arrInput[i].ToString("X2"));
            }
            return sOutput.ToString();
        }

        byte[] GetByteArrayFromLoginAndPassword(string login, string password)
        {
            string sSourceData = login + password;
            byte[] tmpSource = ASCIIEncoding.ASCII.GetBytes(sSourceData);
            byte[] tmpHash = new MD5CryptoServiceProvider().ComputeHash(tmpSource);
            return tmpHash;
        }

        public string GetStringHashSum(string login, string password)
        {
            string str = ByteArrayToString(GetByteArrayFromLoginAndPassword(login, password));
            return str;
        }

        string CreateOracleConnectionString()
        {
            OracleConnectionStringBuilder connectionStringBuilder = new OracleConnectionStringBuilder();
            connectionStringBuilder.DataSource = "localhost:1521/xe";
            connectionStringBuilder.UserID = "system";
            connectionStringBuilder.Password = "123456Qwer";
            return connectionStringBuilder.ConnectionString;
        }

        public void SetOracleConnectionString()
        {
            ConnString = CreateOracleConnectionString();
        }

        partial class INCDOCSDataTable
        {
            public void Fill(OracleConnection OraConnection)
            {
                using (DataSet1TableAdapters.INCDOCSTableAdapter ta = new DataSet1TableAdapters.INCDOCSTableAdapter())
                {
                    ta.Fill(this);
                }
            }

            public void Update(OracleConnection OraConnection)
            {
                using (DataSet1TableAdapters.INCDOCSTableAdapter ta = new DataSet1TableAdapters.INCDOCSTableAdapter())
                {
                    ta.Update(this);
                }
            }
        }


        partial class STAFFDataTable
        {
            public void FindSotr(OracleConnection OraConnection, string emnum)
            {
                using (DataSet1TableAdapters.STAFFTableAdapter ta = new DataSet1TableAdapters.STAFFTableAdapter())
                {
                    ta.FillBy(this, emnum);
                }
            }

            public void FindDirector(OracleConnection OraConnection)
            {
                using (DataSet1TableAdapters.STAFFTableAdapter ta = new DataSet1TableAdapters.STAFFTableAdapter())
                {
                    ta.FindDirector(this);
                }
            }

            public void FindGlavBuh(OracleConnection OraConnection)
            {
                using (DataSet1TableAdapters.STAFFTableAdapter ta = new DataSet1TableAdapters.STAFFTableAdapter())
                {
                    ta.GlavBuh(this);
                }
            }
        }



        partial class REPORTSDataTable
        {
            public void Fill(OracleConnection OraConnection)
            {
                using (DataSet1TableAdapters.REPORTSTableAdapter ta = new DataSet1TableAdapters.REPORTSTableAdapter())
                {
                    ta.Fill(this);
                }
            }
        }

        partial class USERSDataTable
        {
            public bool FindByHash(OracleConnection OraConnection, string hashSum)
            {
                using (DataSet1TableAdapters.USERSTableAdapter ta = new DataSet1TableAdapters.USERSTableAdapter())
                {
                    switch (ta.FindByHash(hashSum))
                    {
                        case 1:
                            return true;
                        default:
                            return false;
                    }
                }
            }

            public void Fill(OracleConnection OraConnection)
            {
                using (DataSet1TableAdapters.USERSTableAdapter ta = new DataSet1TableAdapters.USERSTableAdapter())
                {
                    ta.Fill(this);
                }
            }

            public void Save(OracleConnection OraConnection)
            {
                using (DataSet1TableAdapters.USERSTableAdapter ta = new DataSet1TableAdapters.USERSTableAdapter())
                {
                    ta.Update(this);
                }
            }
        }

    }
}


using System;
using System.Data.OleDb;
using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOM2AMO
{
    public static class DataTypeHelper
    {
        public static OleDbType ToOleDbType(TOM.DataType TOMDataType)
        {
            try
            {
                switch (TOMDataType)
                {
                    case TOM.DataType.Automatic:
                    case TOM.DataType.String:
                        return OleDbType.WChar;
                    case TOM.DataType.Binary:
                        return OleDbType.Binary;
                    case TOM.DataType.Boolean:
                        return OleDbType.Boolean;
                    case TOM.DataType.DateTime:
                        return OleDbType.Date;
                    case TOM.DataType.Decimal:
                        return OleDbType.Decimal;
                    case TOM.DataType.Double:
                        return OleDbType.Double;
                    case TOM.DataType.Int64:
                        return OleDbType.BigInt;
                    case TOM.DataType.Variant:
                        return OleDbType.Variant;
                    default:
                        throw new Exception("The following type is unhandled: " + TOMDataType.ToString());
                }
            }
            catch(Exception e)
            {
                throw new Exception("Error mapping TOM DataType: " + e.Message);
            }
        }
        public static TOM.DataType ToTOMDataType(OleDbType AMODataType)
        {
            try
            {
                switch (AMODataType)
                {
                    case OleDbType.WChar:
                        return TOM.DataType.String;
                    case OleDbType.Binary:
                        return TOM.DataType.Binary;
                    case OleDbType.Boolean:
                        return TOM.DataType.Boolean;
                    case OleDbType.Date:
                        return TOM.DataType.DateTime;
                    case OleDbType.Decimal:
                        return TOM.DataType.Decimal;
                    case OleDbType.Double:
                        return TOM.DataType.Double;
                    case OleDbType.BigInt:
                        return TOM.DataType.Int64;
                    case OleDbType.Variant:
                        return TOM.DataType.Variant;
                    case OleDbType.Empty:
                        return TOM.DataType.Unknown;
                    default:
                        throw new Exception("The following type is unhandled: " + AMODataType.ToString());
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error mapping AMO DataType: " + e.Message);
            }
        }
    }
}

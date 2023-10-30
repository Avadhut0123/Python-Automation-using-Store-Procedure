import pyodbc
import pandas as pd


class Report():

    global cur,OWNER,WAREHOUSE_NAME,SKU,Date
    OWNER = 'WAHL'
    WAREHOUSE_NAME = 'Sai Dhara Bhiwandi'
    SKU = ['WPHS6-0024']
    Date = '2021-12-31'

    conn = ("""driver={SQL Server};server={{Sql Port}};database={{SQL DB_NAME}};
        trusted_connection=no;UID={{SQL UID}};PWD={{SQL PASS}};IntegratedSecurity = true;""")
    conx = pyodbc.connect(conn)
    cur = conx.cursor()

    def Sku_Report(self):

        sql_query = "SELECT OWH.OwnerWarehouseId FROM Owners as ow INNER JOIN OwnerWarehouses AS OWH ON ow.OwnerId = OWH.OwnerId INNER JOIN Warehouses AS WH ON OWH.WarehouseId = WH.WarehouseId WHERE ow.Name = '"+OWNER+"' AND WH.Name = '"+WAREHOUSE_NAME+"'"
        execute_sql = cur.execute(sql_query)
        data = execute_sql.fetchall()
        
        if not data:
            print("Owner/Warehouse Not Available")
        else:
            OwnerWarehouseID = data[0][0]
            for Sku_Name in SKU:
                Sku_Name = Sku_Name

                report_query = pd.read_sql_query("SELECT so.OwnerWarehouseId,so.SoDate,so.SoId,so.ReferenceId,CONVERT(varchar, so.CreatedDate, 106) as SO_Date,som.Description as [Status],sol.SORecordNumber,sOL.ItemId,sol.PiecesOrdered, CASE WHEN sol.PiecesPicked is NULL THEN 0 ELSE sol.PiecesPicked END as PiecesPicked FROM [EMIZA-AUGMENTED-TEST].[dbo].SalesOrderLines sol INNER JOIN [EMIZA-AUGMENTED-TEST].[dbo].SalesOrder so ON sol.SORecordNumber = so.RecordNumber INNER JOIN [EMIZA-AUGMENTED-TEST].[dbo].SOStatusMaster som ON so.StatusId = som.StatusId WHERE SOL.ShowFlag = 'Y'AND OwnerWarehouseId = '"+str(OwnerWarehouseID)+"' AND sol.ItemId IN ('"+str(Sku_Name)+"') AND so.CreatedDate > '"+str(Date)+"'-- AND so.CreatedDate > '2022-09-01'--AND sol.PiecesOrdered <> PiecesPickedorder by sol.SORecordNumber;",self.conx)
                df = pd.DataFrame(report_query)
                df.style
                Report = df.to_csv(r'C:\Users\Emiza\Desktop\ReportSku.csv' ,index= False)
                if Report == False:
                    print("Data Not Available")
                else:
                    print("Report export complete")   

a1 = Report()
a1.Sku_Report()

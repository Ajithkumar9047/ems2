from Config import *
from ExcelContent import create_excel
from EmailContent import email_send
from SqlContent import establish_connection
from test import run
ongo=True
apptry=0
try:
    while ongo:
        rows_query1 = establish_connection(query1)
        num_rows_query1 = len(rows_query1)
        print(f"LeadRepush: {num_rows_query1}")
        logging.info(f"LeadRepush: {num_rows_query1}")

        if num_rows_query1 == 0 :
            rows_query2 = establish_connection( query2)
            print(len(rows_query2))
            rows_query3 = establish_connection( query3)
            print(len(rows_query3))
            exl=create_excel(rows_query2,rows_query3)
            email_send(exl)
            ongo=False
        else:
            print('count is not equal to zero')
            logging.error(f"count is not equal to zero so try reupdate count{apptry+1}")
            run()
            apptry+=1
            if  apptry==3:
                ongo=False

except pyodbc.Error as e:
    logging.error("A pyodbc error occurred: %s", str(e))
    print("A pyodbc error occurred:", str(e))
except Exception as e:
    logging.error("An error occurred: %s", str(e))
    print("A pyodbc error occurred:", str(e))



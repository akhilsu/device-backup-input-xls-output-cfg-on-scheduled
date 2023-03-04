from netmiko import ConnectHandler
from netmiko import NetMikoAuthenticationException, NetMikoTimeoutException
import xlrd
import xlwt
import os
from datetime import datetime
import schedule, time

def doJob():
    print("PROCESS INITIATED..")

    openfile = xlrd.open_workbook(r"device_list.xls")
    sheet = openfile.sheet_by_name("Sheet1")

    print("DATA COPIED SUCCESSFULLY")

    timenow = datetime.now()

    newpath = r"Output\B_"+str(timenow.year)+"-"+str(timenow.month)+"-"+str(timenow.day)+"-"+str(timenow.hour)+"-"+str(timenow.minute)+"-"+str(timenow.second)
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    for i in range (1,sheet.nrows):
        device = {
            "device_type" : sheet.row(i)[5].value,
            "ip" : sheet.row(i)[2].value,
            "username" : sheet.row(i)[3].value,
            "password" : sheet.row(i)[4].value,
            "port" : 22
        }
        print("DEVICE "+ str(int(sheet.row(i)[0].value))+"/"+ str(sheet.nrows-1) +" FETCHED")

        try:
            connectdevice = ConnectHandler(**device)
            print("CONNECTED TO DEVICE")
            hostname = connectdevice.find_prompt()
            fetch = connectdevice.send_command("show running-config", read_timeout=300)

            new_file_create = open(r"Output\B_"+str(timenow.year)+"-"+str(timenow.month)+"-"+str(timenow.day)+"-"+str(timenow.hour)+"-"+str(timenow.minute)+"-"+str(timenow.second)+"\Backup_"+sheet.row(i)[1].value+"-"+"-"+str(timenow.year)+"-"+str(timenow.month)+"-"+str(timenow.day)+"-"+str(timenow.hour)+"-"+str(timenow.minute)+"-"+str(timenow.second)+".cfg","w")
            copy_content = new_file_create.write(fetch)
            new_file_create.close()
            print("BACKUP GENERATED SUCCESSFULLY")
            connectdevice.disconnect()
            print("DISCONNECTED FROM DEVICE\n")
        except NetMikoTimeoutException:
            print("NetMikoTimeoutException - "+sheet.row(i)[1].value)
            new_text_file_create = open(r"Output\B_"+str(timenow.year)+"-"+str(timenow.month)+"-"+str(timenow.day)+"-"+str(timenow.hour)+"-"+str(timenow.minute)+"-"+str(timenow.second)+"\Devices_failed.txt","a")
            copy_content = new_text_file_create.write(sheet.row(i)[1].value+" - NetMikoTimeoutException\n")
            new_text_file_create.close()
        except NetMikoAuthenticationException:
            print("NetMikoAuthenticationException - "+sheet.row(i)[1].value)
            new_text_file_create = open(r"Output\B_"+str(timenow.year)+"-"+str(timenow.month)+"-"+str(timenow.day)+"-"+str(timenow.hour)+"-"+str(timenow.minute)+"-"+str(timenow.second)+"\Devices_failed.txt","a")
            copy_content = new_text_file_create.write(sheet.row(i)[1].value+" - NetMikoAuthenticationException\n")
            new_text_file_create.close()

    print("PROCESS FINISHED SUCCESSFULLY " + str(datetime.now()) + ", It took " + str(datetime.now()-timenow))

doJob()

schedule.every(3600).seconds.do(doJob)


while True:
    schedule.run_pending()
    time.sleep(3600)


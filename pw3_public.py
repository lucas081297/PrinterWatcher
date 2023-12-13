import time
import win32print
import pywintypes
import requests
import emoji
import datetime
import keyboard

#----------------------------------------------------------------------

# Funcao de envio de mensagem via API Telegram
def send_message(token, chat_id, message):
    try:
        data = {"chat_id": chat_id, "text": message}
        url = "https://api.telegram.org/bot{}/sendMessage".format(token)
        requests.post(url, data)
    except Exception as e:
        print("Error in sendMessage:", e)

def del_job(phandle, jobid):
    try: 
        win32print.SetJob(phandle, jobid, 0, None, win32print.JOB_CONTROL_CANCEL)
    except pywintypes.error as e:
        if e.winerror == 5:
            print("Permission Error")
        else:
            raise e

#Variables
chat_id = '' #Place your chat_id group Telegram
token = '' #Place your token Telegram
data = ""


# Core function
def print_job_checker():
    """
    """
    data = ""
    # While jobs...
    while True:
        try:
            # Browse all printers installed on the local machine
            for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
                flags, desc, name, comment = p
                # Creates a handle opening the printer from the name obtained previously
                phandle = win32print.OpenPrinter(name, {'DesiredAccess': win32print.PRINTER_ACCESS_USE})
                # List the entire printer print queue
                print_jobs = win32print.EnumJobs(phandle, 0, -1, 1)
                # Check if there is any job in the print queue, if yes, check the job
                if print_jobs:
                    jobs_with_error = []
                    for job in print_jobs:
                        jobid = job["JobId"]
                        document = job["pDocument"]
                        nick = job["pUserName"]
                        status = job["Status"]
                        # Check job flag JOB_STATUS_ERROR
                        if status & win32print.JOB_STATUS_ERROR:
                            # Adds the job with error to the list of jobs with error
                            jobs_with_error.append(jobid)
                            data = f"{emoji.emojize(':warning:')}JOB WITH ERROR!{emoji.emojize(':warning:')}\n\n{emoji.emojize(':printer:')}Printer: {str(name)}\n\n"
                            data += f"{emoji.emojize(':spiral_notepad:')}File : {str(document)}\n{emoji.emojize(':person_bowing:')}User: {str(nick)}\n\n"
                    if jobs_with_error:
                        # Cancel jobs with errors
                        for jobid in jobs_with_error:
                            win32print.SetJob(phandle, jobid, 0, None, win32print.JOB_CONTROL_CANCEL)
                        # Calls the message sending function via Telegram    
                        send_message(token, chat_id, data)
                # Close the Printer Handle
                win32print.ClosePrinter(phandle)

            if data:
                print(data)
            print(f"\nFinalizado às {datetime.datetime.now()}\n")
            time.sleep(30)  
            
            print ("Fila vazia!")
            time.sleep(30)
        
        except NameError:
            print(NameError)



print("PrinterWatcher 1.1\nStart...")
print("To stop press: ctrl + c") 
while keyboard.is_pressed('esc') == False:
    try:
        print_job_checker()
    except NameError:
        pass

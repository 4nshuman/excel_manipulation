import Work_order as wo, incidents as inci, root_cause_analysis as rca
import logging, datetime, os
LOG_FILENAME = 'SCRIPT.log'

src = raw_input('Enter the source file name (with extension) : ')
dir_name = os.getcwd()

try:
    wo.ground_zero_wo(src,dir_name)
    inci.ground_zero_i(src,dir_name)
    rca.ground_zero_rc(src,dir_name)
    print("All sheets written successfully")
    logging.basicConfig(filename=LOG_FILENAME,level=logging.INFO)
    logging.info("All sheets written successfully at "+str(datetime.datetime.now()))
except Exception as e:
    logging.basicConfig(filename=LOG_FILENAME,level=logging.ERROR)
    logging.error("\n"+str(e)+" : "+str(datetime.datetime.now()))
    print("Oops something went wrong ! ")
    print("I suppose you entered a wrong file name.")
    print("Check if you added the extension.")

# excel_manipulation
reading and editing excel files with macros using python

1. get_file_data.py --
        This file is used to get sheet objects. It returns sheets from where to read the data. Uses xlrd.
2. incidents.py --
        This module is used to calculate the various requirements of the excel file.
3. incidents_write.py --
        This file is used to write the data calculated by the incidents.py into the file. Uses openpyxl.
4. root_cause_analysis --
        This module is used to perform the analysis of various problem tickets and find out their root cause. This module also writes in the data to the excel file. Uses openpyxl for writing purposes and xlrd for reading.
5. start_point.py --
        This file calls various modules to perform the complete task.
6. work_order.py --
        This module is used to calculate as well as write the work order data. It reads from the excel and makes use of xlrd for the same and writes in the file using openpyxl.
7. SCRIPT.log --
        Stores the complete log data for the modules.
8. repo.xlsm --
        This xlsm file contains the source data and in this file itself the calculated data is inserted. Note that this file makes use of various macros to prepare charts, and format cells.

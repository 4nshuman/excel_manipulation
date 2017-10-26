from __future__ import print_function
import xlrd
import get_file_data as gfd, incidents_write as iW

def generate_column_indices(cal_sheet):
    global priority_col_idx
    global time_col_idx
    global status_col_idx
    global app_col_idx

    # Print all values, iterating through rows and columns
    num_cols = cal_sheet.ncols   # Number of columns
    for col_idx in range(0, num_cols):  # Iterate through columns
        cell_obj = cal_sheet.cell(0, col_idx)  # Get cell object by row, col
        if(cell_obj.value == 'Internal Priority'):
            priority_col_idx=col_idx
        if(cell_obj.value == 'Time'):
            time_col_idx=col_idx
        if(cell_obj.value == 'Status'):
            status_col_idx=col_idx
        if(cell_obj.value == 'Application Name'):
            app_col_idx=col_idx

def calculations(cal_sheet):
    #GETTING TICKET COUNT
    global total_tickets,total_p1_p2,total_time
    total_tickets=cal_sheet.nrows-1
    global ecc_p1_p2_count,ecc_total_count,filenet_p1_p2_count,filenet_total_count
    global seds_p1_p2_count,seds_total_count,kofax_p1_p2_count,kofax_total_count
    global ilinx_p1_p2_count,ilinx_total_count,alexandria_p1_p2_count,alexandria_total_count
    ecc_p1_p2_count=ecc_total_count=filenet_p1_p2_count=filenet_total_count=0
    seds_p1_p2_count=seds_total_count=kofax_p1_p2_count=kofax_total_count=0
    ilinx_p1_p2_count=ilinx_total_count=alexandria_p1_p2_count=alexandria_total_count=0
    for row_idx in range(1, cal_sheet.nrows):# Iterate through rows
        cell_obj = cal_sheet.cell(row_idx, priority_col_idx)  # Get cell object by row, col
        if(cell_obj.value==2 or cell_obj.value==1):
            if(cal_sheet.cell(row_idx, app_col_idx).value=='ECC'):
                ecc_p1_p2_count+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='SEDS'):
                seds_p1_p2_count+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='FileNet'):
                filenet_p1_p2_count+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='Kofax'):
                kofax_p1_p2_count+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='ILINX'):
                ilinx_p1_p2_count+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='Alexandria'):
                alexandria_p1_p2_count+=1
        if(cal_sheet.cell(row_idx, app_col_idx).value=='ECC'):
            ecc_total_count+=1
        elif(cal_sheet.cell(row_idx, app_col_idx).value=='SEDS'):
            seds_total_count+=1
        elif(cal_sheet.cell(row_idx, app_col_idx).value=='FileNet'):
            filenet_total_count+=1
        elif(cal_sheet.cell(row_idx, app_col_idx).value=='Kofax'):
            kofax_total_count+=1
        elif(cal_sheet.cell(row_idx, app_col_idx).value=='ILINX'):
            ilinx_total_count+=1
        elif(cal_sheet.cell(row_idx, app_col_idx).value=='Alexandria'):
            alexandria_total_count+=1
    total_p1_p2 = ecc_p1_p2_count+seds_p1_p2_count+filenet_p1_p2_count+kofax_p1_p2_count+ilinx_p1_p2_count

    #GETTING RESOLVED OPEN AND CLOSED COUNT
    global ecc_closed, ecc_resolved, ecc_open, seds_closed, seds_resolved, seds_open, filenet_closed, filenet_resolved, filenet_open
    global kofax_closed, kofax_resolved, kofax_open, ilinx_closed, ilinx_resolved, ilinx_open, alexandria_closed, alexandria_resolved, alexandria_open
    ecc_closed=ecc_resolved=ecc_open=seds_closed=seds_resolved=seds_open=filenet_closed=filenet_resolved=filenet_open = 0
    kofax_closed=kofax_resolved=kofax_open=ilinx_closed=ilinx_resolved=ilinx_open=alexandria_closed=alexandria_resolved=alexandria_open = 0
    for row_idx in range(1, cal_sheet.nrows):# Iterate through rows
        cell_obj = cal_sheet.cell(row_idx, status_col_idx)  # Get cell object by row, col
        if(cell_obj.value=='CLOSED'):
            if(cal_sheet.cell(row_idx, app_col_idx).value=='ECC'):
                ecc_closed+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='SEDS'):
                seds_closed+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='FileNet'):
                filenet_closed+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='Kofax'):
                kofax_closed+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='ILINX'):
                ilinx_closed+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='Alexandria'):
                alexandria_closed+=1
        elif(cell_obj.value=='OPEN'):
            if(cal_sheet.cell(row_idx, app_col_idx).value=='ECC'):
                ecc_open+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='SEDS'):
                seds_open+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='FileNet'):
                filenet_open+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='Kofax'):
                kofax_open+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='ILINX'):
                ilinx_open+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='Alexandria'):
                alexandria_open+=1
        elif(cell_obj.value=='RESOLVED'):
            if(cal_sheet.cell(row_idx, app_col_idx).value=='ECC'):
                ecc_resolved+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='SEDS'):
                seds_resolved+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='FileNet'):
                filenet_resolved+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='Kofax'):
                kofax_resolved+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='ILINX'):
                ilinx_resolved+=1
            elif(cal_sheet.cell(row_idx, app_col_idx).value=='Alexandria'):
                alexandria_resolved+=1

    # GETTING TOTAL TIME
    global ecc_time,kofax_time,ilinx_time,filenet_time,seds_time,alexandria_time
    ecc_time=kofax_time=ilinx_time=filenet_time=seds_time=alexandria_time=0
    for row_idx in range(1, cal_sheet.nrows):    # Iterate through rows
        cell_obj = cal_sheet.cell(row_idx, time_col_idx)  # Get cell object by row, col
        if(cal_sheet.cell(row_idx, priority_col_idx).value==1 or cal_sheet.cell(row_idx, priority_col_idx).value==2):
            if(cal_sheet.cell(row_idx, app_col_idx).value=='ECC'):
                ecc_time+=float(cell_obj.value)
            if(cal_sheet.cell(row_idx, app_col_idx).value=='SEDS'):
                seds_time+=float(cell_obj.value)
            if(cal_sheet.cell(row_idx, app_col_idx).value=='FileNet'):
                filenet_time+=float(cell_obj.value)
            if(cal_sheet.cell(row_idx, app_col_idx).value=='Kofax'):
                kofax_time+=float(cell_obj.value)
            if(cal_sheet.cell(row_idx, app_col_idx).value=='ILINX'):
                ilinx_time+=float(cell_obj.value)
            if(cal_sheet.cell(row_idx, app_col_idx).value=='Alexandria'):
                alexandria_time+=float(cell_obj.value)
    total_time=ecc_time+seds_time+filenet_time+kofax_time+ilinx_time+alexandria_time

def display():
    print ('ECC P1/P2 Count: [%s]' % ecc_p1_p2_count)
    print ('FileNet P1/P2 Count: [%s]' % filenet_p1_p2_count)
    print ('SEDS Total Count: [%s]' % seds_p1_p2_count)
    print ('KOFAX P1/P2 Count: [%s]' % kofax_p1_p2_count)
    print ('ILINX P1/P2 Count: [%s]' % ilinx_p1_p2_count)

    print ('ECC Count: [%s]' % ecc_total_count)
    print ('FileNet Count: [%s]' % filenet_total_count)
    print ('SEDS Count: [%s]' % seds_total_count)
    print ('KOFAX Count: [%s]' % kofax_total_count)
    print ('ILINX Count: [%s]' % ilinx_total_count)

    print ('Total tickets: [%s]' % total_tickets)
    print ('p1/p2 tickets: [%s]' % total_p1_p2)
    print ('Total Time: [%s]' % total_time)

    print ('RESOLVED SEDS: [%s]' % seds_closed)

def create_dict():
    global incident_data_dict
    incident_data_dict={}
    incident_data_dict['ecc_p1_p2_count'] =  ecc_p1_p2_count
    incident_data_dict['filenet_p1_p2_count'] = filenet_p1_p2_count
    incident_data_dict['seds_p1_p2_count'] = seds_p1_p2_count
    incident_data_dict['kofax_p1_p2_count'] = kofax_p1_p2_count
    incident_data_dict['ilinx_p1_p2_count'] = ilinx_p1_p2_count
    incident_data_dict['alexandria_p1_p2_count'] = alexandria_p1_p2_count
    incident_data_dict['total_p1_p2'] = total_p1_p2

    incident_data_dict['ecc_total_count'] = ecc_total_count
    incident_data_dict['filenet_total_count'] = filenet_total_count
    incident_data_dict['seds_total_count'] = seds_total_count
    incident_data_dict['kofax_total_count'] = kofax_total_count
    incident_data_dict['ilinx_total_count'] = ilinx_total_count
    incident_data_dict['alexandria_total_count'] = alexandria_total_count
    incident_data_dict['total_tickets'] = total_tickets

    incident_data_dict['ecc_time'] = ecc_time
    incident_data_dict['kofax_time'] = kofax_time
    incident_data_dict['ilinx_time'] = ilinx_time
    incident_data_dict['filenet_time'] = filenet_time
    incident_data_dict['alexandria_time'] = alexandria_time
    incident_data_dict['seds_time'] = seds_time
    incident_data_dict['total_time'] = total_time

    incident_data_dict['ecc_closed']=ecc_closed
    incident_data_dict['ecc_resolved']=ecc_resolved
    incident_data_dict['ecc_open']=ecc_open
    incident_data_dict['seds_closed']=seds_closed
    incident_data_dict['seds_resolved']=seds_resolved
    incident_data_dict['seds_open']=seds_open
    incident_data_dict['filenet_closed']=filenet_closed
    incident_data_dict['filenet_resolved']=filenet_resolved
    incident_data_dict['filenet_open']=filenet_open
    incident_data_dict['kofax_closed']=kofax_closed
    incident_data_dict['kofax_resolved']=kofax_resolved
    incident_data_dict['kofax_open']=kofax_open
    incident_data_dict['ilinx_closed']=ilinx_closed
    incident_data_dict['ilinx_resolved']=ilinx_resolved
    incident_data_dict['ilinx_open']=ilinx_open
    incident_data_dict['alexandria_closed']=alexandria_closed
    incident_data_dict['alexandria_resolved']=alexandria_resolved
    incident_data_dict['alexandria_open']=alexandria_open

def ground_zero_i(src,dir_name):
    sheet = gfd.get_sheet_readable(current_dir=dir_name, filename=src, sheet_index=3)
    generate_column_indices(sheet)
    calculations(sheet)
    create_dict()
    iW.incident_write(src,dir_name,incident_data_dict)

print('''
This script is meant to amalgamate multiple excel reports in
such a way to make change detection easy for the end user.

Instruction on formatting your input excel files:

    For the best results, ensure that the first column of each
    of your excel files is the ID colum of your objects. 

    If it is not, you will need to modify the excel files and
    put the ID column as the first column ("A") of your excel
    sheet.

    Aslo please make sure that all data is in the very first
    sheet (or the only sheet) of your excel file. The rest of
    the sheets will be ignored by the script in its current
    setup.

Limitation:

    On my system, the maximum size of the log files can be
    somewhere between 32000 and 33000 Kb. For this reason, I
    made the Process Log file smaller by breaking the log
    in 2 halves.

    Please be advised that not all systems may have the same
    limitations. If you feel your log files require more space,
    then the main log dictionary should be divided further and
    additional log files created.

    Author:         Seyed-Mahdy Sadraddini
----------------------------------------------------------------
''')

import pandas as pd
import os, numpy, time
from datetime import datetime

starttime = time.time()
timestamp = time.asctime()
    # this is the identifying timestamp for the same bunch of files
print(timestamp+'\n\n')

all_ID = dict()

files = filter(os.path.isfile, os.listdir(os.getcwd()))
files = [os.path.join(os.getcwd(), File) for File in files]
files.sort(key = lambda ts: os.path.getmtime(ts))

if not os.path.exists(os.path.join(os.getcwd(),\
    os.path.basename(files[0]).split('.')[0])):
    os.mkdir(os.path.join(os.getcwd(),\
        os.path.basename(files[0]).split('.')[0]))
    print(os.path.basename(files[0]).split('.')[0]+\
          ' folder is created on the main files directory...\n\n')

d = {}

with open(os.path.join(os.path.join(os.getcwd(),\
    os.path.basename(files[0]).split('.')[0]),\
    os.path.basename(files[0]).split('.')[0]+'_log.txt'),\
    'w') as logger:

    print('Primary Logger file is created...')
    
    logger.write('EXCEL AMALGAMATION LOGGER\n')
    logger.write('-'*50)
    logger.write('\n')

    logger.write(timestamp)
    logger.write('\n')
    logger.write('-'*50)
    logger.write('\n')

    d['start_datetime'] = timestamp

    logger.write('ORDER OF FILES BASED ON LAST MODIFIED DATE:\n')
    logger.write('-'*50)
    logger.write('\n\n')
    for f in files:
        logger.write(f)
        logger.write('\n')

    logger.write('\n\n')
    id_list = list() # generate a list of IDs for each file

    if not os.path.exists(os.path.join(os.path.join(os.getcwd(),\
        os.path.basename(files[0]).split('.')[0]),'FileErrorLog')):
        os.mkdir(os.path.join(os.path.join(os.getcwd(),\
            os.path.basename(files[0]).split('.')[0]),'FileErrorLog'))
        print('\n\nFileErrorLog folder is created...\n\n')

    for i in files:
        print('Reading File {}'.format(i))
        
        if '.py' in i: pass
        else:
            logger.write(i)
            logger.write('\n')
            logger.write('-'*50)
            logger.write('\n')
            logger.write('reading the file into a dataframe...')
            unpk = pd.read_excel(i)
            logger.write('\n')
            logger.write('\n')

            with open(os.path.join(os.path.join(os.path.join(os.getcwd(),\
                os.path.basename(files[0]).split('.')[0]),'FileErrorLog'),\
                os.path.basename(i).split('.')[0]+'_ErrorLog.txt'),'w') as fileErrorLogger:

                fileErrorLogger.write('Error Log')
                fileErrorLogger.write('\n')
                fileErrorLogger.write('-'*50)
                fileErrorLogger.write(timestamp)
                fileErrorLogger.write('\n\n')
                fileErrorLogger.write(str(i))
                fileErrorLogger.write('\n')
                fileErrorLogger.write('-'*50)
                fileErrorLogger.write('\n')

                count_error = 0
                for index, row in unpk.iterrows():
                    id_list.append(row[0])
                    
                    try:
                        ID = row[0].decode().strip()
                        if not ID in d:
                            d[ID] = [row[1:]]
                        else:
                            d[ID].append(row[1:])
                        logger.write(str(row[0]))
                        logger.write('\n')

                    except:
                        count_error+=1
                        fileErrorLogger.write(str(ID))
                        fileErrorLogger.write('\n')
                        pass

                fileErrorLogger.write('\n\nNumber of errors logged for this file:')
                fileErrorLogger.write('\t'+str(count_error))
                print('File Error Log was created for {}'.\
                      format(str(os.path.basename(i))))

                logger.write('\n')
                logger.write('-'*50)
                logger.write('\n')
                
            logger.write('\n')
            logger.write(str(str(len(id_list))+' IDs'))
            logger.write('\n\n')
            for item in id_list:
                if item not in all_ID:
                    all_ID[item] = 1
                else:
                    all_ID[item] +=1
                    
            id_list = []

    logger.write('\n\nWRITING RESULTS TO PROCESS LOG...\n')
    logger.write('-'*50)
    logger.write('\n')

    print('\n\nProcess Log Files creation in progress...\n\n')
    # Breaking down the dictionary in two halves
    d1 = dict(d.items()[:len(d)/2])
    d2 = dict(d.items()[len(d)/2:])

    f1 = open(os.path.join(os.path.join(os.getcwd(),\
    os.path.basename(files[0]).split('.')[0]),\
    os.path.basename(files[0]).split('.')[0]+'_ProcessLog1.txt'),\
    'w')
    f2 = open(os.path.join(os.path.join(os.getcwd(),\
    os.path.basename(files[0]).split('.')[0]),\
    os.path.basename(files[0]).split('.')[0]+'_ProcessLog2.txt'),\
    'w')

    c1 = 0
    for k1, v1 in d1.items():
        c1 +=1
        f1.write(k1)
        f1.write('\n')
        for val in v1:
            f1.write(str(val))
            f1.write('\n')
        f1.write('-'*75)
        f1.write('\n')
    f1.write('\n\n\nNumber of IDs in this file:\n\t')
    f1.write(str(c1))
    f1.write('\n\n\n')
    f1.write(timestamp)
    f1.close()

    c2 = 0
    for k2, v2 in d2.items():
        c2 +=1
        f2.write(k2)
        f2.write('\n')
        for val2 in v2:
            f2.write(str(val2))
            f2.write('\n')
        f2.write('-'*75)
        f2.write('\n')
    f2.write('\n\n\nNumber of IDs in this file:\n\t')
    f2.write(str(c2))
    f2.write('\n\n\n')
    f2.write(timestamp)
    f2.close()
                
    logger.write(str(str(len(d))+\
        ' keys were written to the process log file.'))

print('Creating Log for present IDs and their frequency in different files.\n')

with open(os.path.join(os.path.join(os.getcwd(),\
    os.path.basename(files[0]).split('.')[0]),'all_ID Occurrence.txt'),'w')\
    as countID:
    countID.write('Number of Times the IDs have Occurred in Different Files')
    countID.write('\n')
    countID.write('-'*50)
    countID.write(timestamp)
    countID.write('\n\n')
    for k, v in all_ID.items():
        countID.write(str('{0}\t\t{1}'.format(k,v)))
        countID.write('\n')
    countID.write('\n\n\n')
    countID.write('Number of Distinct IDs present:\n')
    countID.write(str(len(all_ID)))

endtime = time.time()
timespan = endtime - starttime
print('The amalgamation process took '+str(timespan)+' seconds.')

time.sleep(2)

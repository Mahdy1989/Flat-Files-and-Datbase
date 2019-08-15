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

    Sometimes the writer might arbitrarily exclude some information
    by setting an arbitrary limit on the size of the text.
    I have not found a solution to this problem as I am not able to
    replicate this problem. For this reason, I figured a way around
    it: I made the Process Log file smaller by breaking the log in
    2 halves. To this end, please use the following script:

                        amalgamateExcels_v2.py

    It is necessary that you check all the log files to ensure the
    amalgamation prcess is excuted successfully. I have put in place
    some error handling logic to log possible errors. Therefore you
    should not avoid this step at any cost! Checking logs is a must;
    especially contents of FileErrorLog folder and comparing the
    number of output IDs in:
                        "all_ID Changes.txt"
                                &
                        "..._ProcessLog.txt"

    Author:         Seyed-Mahdy Sadraddini
----------------------------------------------------------------
''')


import pandas as pd
import os, numpy, time, shutil
from datetime import datetime

starttime = time.time()
timestamp = time.asctime()
    # this is the identifying timestamp for the same bunch of files

all_ID = dict()

files = filter(os.path.isfile, os.listdir(os.getcwd()))
files = [os.path.join(os.getcwd(), File) for File in files]
files.sort(key = lambda ts: os.path.getmtime(ts))

if not os.path.exists(os.path.join(os.getcwd(),\
    os.path.basename(files[0]).split('.')[0])):
    os.mkdir(os.path.join(os.getcwd(),\
        os.path.basename(files[0]).split('.')[0]))

d = {}

with open(os.path.join(os.path.join(os.getcwd(),\
    os.path.basename(files[0]).split('.')[0]),\
    os.path.basename(files[0]).split('.')[0]+'_log.txt'),\
    'w') as logger:
    # This is the main logger for the script ..._log.txt
    
    logger.write('EXCEL AMALGAMATION LOGGER\n')
    logger.write('-'*50)
    logger.write('\n')

    print(timestamp)
    logger.write(timestamp)
    logger.write('\n')
    logger.write('-'*50)
    logger.write('\n')

    d['start_datetime'] = timestamp

    logger.write('\n\nORDER OF FILES BASED ON LAST MODIFIED DATE:\n')
    logger.write('-'*50)
    for f in files:
        logger.write(f)
        logger.write('\n')

    logger.write('\n\n')
    id_list = list() # generate a list of IDs for each file
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

            if not os.path.exists(os.path.join(os.path.join(os.getcwd(),\
                os.path.basename(files[0]).split('.')[0]),'FileErrorLog')):
                os.mkdir(os.path.join(os.path.join(os.getcwd(),\
                    os.path.basename(files[0]).split('.')[0]),'FileErrorLog'))
                    # FileErrorLog folder
                print('\n\nFileErrorLog folder is created...\n\n')
            with open(os.path.join(os.path.join(os.path.join(os.getcwd(),\
                os.path.basename(files[0]).split('.')[0]),'FileErrorLog'),\
                os.path.basename(i).split('.')[0]+'_ErrorLog.txt'),'w') as fileErrorLogger:
                # FileErrorLogger

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

                    except Exception as e:
                        count_error+=1
                        fileErrorLogger.write(str(ID))
                        fileErrorLogger.write('\n')
                        fileErrorLogger.write(str(e))
                        fileErrorLogger.write('\n\n')
                        continue

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

    print('\nWriting the process log...\n...\n...')           
    with open(os.path.join(os.path.join(os.getcwd(),\
        os.path.basename(files[0]).split('.')[0]),\
        os.path.basename(files[0]).split('.')[0]+'_ProcessLog.txt'),\
        'w') as processLog:
        # Process Log
        
        logger.write('\n\nWRITING RESULTS TO PROCESS LOG...\n')
        logger.write('-'*50)
        logger.write('\n')
        
        counter = 0
        num = 1
        for key, value in d.items():
            counter +=1
            
            if key == 'start_datetime':
                pass
            
            processLog.write(key)
            processLog.write('\n')
            
            for val in value:
                processLog.write(str(val)+'\n')
                              
        processLog.write('\n')
        processLog.write('-'*75)
        processLog.write('\n')
            

        processLog.write(timestamp)
        logger.write(str(str(counter-1)+\
        ' keys were written to the process log file.'))

print('Creating Log for present IDs and their frequency in different files.\n')

with open(os.path.join(os.path.join(os.getcwd(),\
    os.path.basename(files[0]).split('.')[0]),'all_ID Changes.txt'),'w') as countID:
    #all_ID Changes.txt
    countID.write('Number of Times the IDs have Occurred in Different Files')
    countID.write('\n')
    countID.write('-'*50)
    countID.write(timestamp)
    countID.write('\n\n')
    for k, v in all_ID.items():
        countID.write(str('{0}\t{1}'.format(k,v)))
        countID.write('\n')
    countID.write('\n\n\n')
    countID.write('Number of Distinct IDs present:\n')
    countID.write(str(len(all_ID)))

endtime = time.time()
timespan = endtime - starttime
print('The amalgamation process took '+str(timespan)+' seconds.')

time.sleep(2)

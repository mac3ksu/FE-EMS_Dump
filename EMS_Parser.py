import os
import shutil
import time
import xlrd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def grab_rtu_list(worksheet):
    rtus_raw = worksheet.col(2)
    rtus = []
    for rtu in rtus_raw:
        if rtu.value != '':
            rtus.append(rtu.value)
    rtus.pop(0)
    rtus.sort()
    #print(rtus)
    return rtus

def archive_fep_files(region, rtus):
    # archive all existing fep/ch/baud dump csv files
    for i, rtu in enumerate(rtus):
        print('archive fep {}/{}'.format(i + 1, len(rtus)))

        # if an archive file is not existing inside the specific rtu folder,
        #          create one for the purpose of archiving legacy EMS dumps
        #          archive_dir_dest: archive directory destination
        archive_dir_dest = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + '_FEP' + '\\' + 'Archive' + '\\'
        if not os.path.exists(archive_dir_dest):
            os.makedirs(archive_dir_dest)

        outfile_dir = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + '_FEP' + '\\'

        # move all old dump worksheets from inside rtu folder to the archive folder. we want to keep them, but it looks cluttered
        #           this direction will execute before the csv files are created and put into the rtu folder
        #           the goal is for all legacy dump files to be archived first allowing for only the new files to remain
        files2move = os.listdir(outfile_dir)
        for file in files2move:
            shutil.move(outfile_dir + file, archive_dir_dest)  # shutil.move(source, destination)

def archive_rtu_files(region, worksheet, rtus):

    rtu_dict = {}
    for rtu in rtus:
        rtu_dict[rtu] = []

    for i, entry in enumerate(worksheet.col(2)):
        if i:
            try:
                rtu_dict[entry.value].append(i)
            except:
                pass

    # for all rtus in the ems dump excel doc, archive existing csv files in specified folder
    for i, rtu in enumerate(rtus):
        print('archive {}/{}'.format(i + 1, len(rtus)))

        # if an archive file is not existing inside the specific rtu folder,
        #          create one for the purpose of archiving legacy EMS dumps
        #          archive_dir_dest: archive directory destination
        archive_dir_dest = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + rtu + '\\' + 'Archive' + '\\'
        if not os.path.exists(archive_dir_dest):
            os.makedirs(archive_dir_dest)

        outfile_dir = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + rtu + '\\'

        # move all old dump worksheets from inside rtu folder to the archive folder. we want to keep them, but it looks cluttered
        #           this direction will execute before the csv files are created and put into the rtu folder
        #           the goal is for all legacy dump files to be archived first allowing for only the new files to remain
        files2move = os.listdir(outfile_dir)
        for file in files2move:
            shutil.move(outfile_dir + file, archive_dir_dest)  # shutil.move(source, destination)

def rtu_fep_parse(region, date, worksheet, rtus):
    rtu_dict = {}
    for rtu in rtus:
        rtu_dict[rtu] = []

    for i, entry in enumerate(worksheet.col(2)):
        if i:
            try:
                rtu_dict[entry.value].append(i)
            except:
                pass

    # for all rtus in the ems dump excel doc, update or add fep/ch/baud csv files in the fep folder
    for i, rtu in enumerate(rtus):
        print('fep/ch/baud {}/{}'.format(i + 1, len(rtus)))

        outfile_dir = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + '_FEP' + '\\'
        if not os.path.exists(outfile_dir):
            os.makedirs(outfile_dir)

        # for the specific rtu, create a fep/ch/baud dump csv document
        outfile_name = rtu + '_' + date + '_FEP.csv'
        outfile = outfile_dir + outfile_name

        # print(rtu_dict[rtu])
        # populate the rtu fep/ch/baud dump csv file with values from ems dump
        with open(outfile, 'w+') as output_file:
            output_file.write(
                'RTUC: 1989, PATH, RTU from RTUC, TYPE_RTU, ADDR_RTU, BAUD_PATH\n')
            for row in rtu_dict[rtu]:
                output_file.write('{},{},{},{},{},{}\n'.format(
                    worksheet.cell(row, 0).value,
                    worksheet.cell(row, 1).value,
                    worksheet.cell(row, 2).value,
                    worksheet.cell(row, 3).value,
                    worksheet.cell(row, 4).value,
                    worksheet.cell(row, 5).value,
                ))

def status_parse(region, date, worksheet, rtus):
    rtu_dict = {}
    for rtu in rtus:
        rtu_dict[rtu] = []

    for i, entry in enumerate(worksheet.col(1)):
        if i:
            try:
                rtu_dict[entry.value].append(i)
            except:
                pass

    # for all rtus in the ems dump excel doc, create a csv file to host the ems dump status points
    for i, rtu in enumerate(rtus):
        print('status {}/{}'.format(i+1, len(rtus)))
        # print(rtu)

        # if an ems dump file for a specific rtu is not existing inside the location folder, create one
        outfile_dir = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + rtu + '\\'
        if not os.path.exists(outfile_dir):
            os.makedirs(outfile_dir)

        # for the specific rtu, create a status dump csv document
        outfile_name = date + '_' + rtu + '_STATUS.csv'
        outfile = outfile_dir + outfile_name

        # print(rtu_dict[rtu])
        # populate the rtu status dump csv file with values from ems dump
        with open(outfile, 'w+') as output_file:
            output_file.write(
                'STATION, RTU, TYPE_RTU, RTU_STATUS, PHYADR, EMS POINT, PRI SITE, SEC SITE2, SINVT, XINVT, MCD, CONCAT, CONV, ID_DEVICE (short), NAME_DEVICE (descriptive)\n')
            for row in rtu_dict[rtu]:
                output_file.write('{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}\n'.format(
                    worksheet.cell(row, 0).value,
                    worksheet.cell(row, 1).value,
                    worksheet.cell(row, 2).value,
                    worksheet.cell(row, 3).value,
                    worksheet.cell(row, 4).value,
                    worksheet.cell(row, 5).value,
                    worksheet.cell(row, 6).value,
                    worksheet.cell(row, 7).value,
                    worksheet.cell(row, 8).value,
                    worksheet.cell(row, 9).value,
                    worksheet.cell(row, 10).value,
                    worksheet.cell(row, 11).value,
                    worksheet.cell(row, 12).value,
                    worksheet.cell(row, 13).value,
                    worksheet.cell(row, 14).value,
                ))
        # print(rtu + ' completed')


def control_parse(region, date, worksheet, rtus):
    rtu_dict = {}
    for rtu in rtus:
        rtu_dict[rtu] = []

    for i, entry in enumerate(worksheet.col(1)):
        if i:
            try:
                rtu_dict[entry.value].append(i)
            except:
                pass

    for i, rtu in enumerate(rtus):
        print('control {}/{}'.format(i + 1, len(rtus)))
        outfile_dir = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + rtu + '\\'
        if not os.path.exists(outfile_dir):
            os.makedirs(outfile_dir)

        outfile_name = date + '_' + rtu + '_CONTROL.csv'
        outfile = outfile_dir + outfile_name

        with open(outfile, 'w+') as output_file:
            output_file.write('STATION,RTU,TYPE_RTU,RTU CONTROL,CONTROL,PHYADR_RELAY,EMS CONTROL,ID_CTRL,CTRLFUNC,COMMAND,SEXP,OPTIME,WAIT,TIMEOUT,ID_DEVICE (short),NAME_DEVICE (descriptive)\n')
            for row in rtu_dict[rtu]:
                output_file.write('{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}\n'.format(
                    worksheet.cell(row, 0).value,
                    worksheet.cell(row, 1).value,
                    worksheet.cell(row, 2).value,
                    worksheet.cell(row, 3).value,
                    worksheet.cell(row, 4).value,
                    worksheet.cell(row, 5).value,
                    worksheet.cell(row, 6).value,
                    worksheet.cell(row, 7).value,
                    worksheet.cell(row, 8).value,
                    worksheet.cell(row, 9).value,
                    worksheet.cell(row, 10).value,
                    worksheet.cell(row, 11).value,
                    worksheet.cell(row, 12).value,
                    worksheet.cell(row, 13).value,
                    worksheet.cell(row, 14).value,
                ))


def analog_parse(region, date, worksheet, rtus):
    rtu_dict = {}
    for rtu in rtus:
        rtu_dict[rtu] = []

    for i, entry in enumerate(worksheet.col(1)):
        if i:
            try:
                rtu_dict[entry.value].append(i)
            except:
                pass

    for i, rtu in enumerate(rtus):
        print('analog {}/{}'.format(i + 1, len(rtus)))
        outfile_dir = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + rtu + '\\'
        if not os.path.exists(outfile_dir):
            os.makedirs(outfile_dir)

        outfile_name = date + '_' + rtu + '_ANALOG.csv'
        outfile = outfile_dir + outfile_name

        with open(outfile, 'w+') as output_file:
            output_file.write('STATION,RTU,TYPE_RTU,RTU ANALOG,PHYADR,EMS ANALOG,PRI SITE,SEC SITE2,loreas,hireas,RAW LOW,RAW HIGH,ENG LOW,ENG HIGH,NEGATE,ID_DEVICE (short),NAME_DEVICE (descriptive)\n')
            for row in rtu_dict[rtu]:
                output_file.write('{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}\n'.format(
                    worksheet.cell(row, 0).value,
                    worksheet.cell(row, 1).value,
                    worksheet.cell(row, 2).value,
                    worksheet.cell(row, 3).value,
                    worksheet.cell(row, 4).value,
                    worksheet.cell(row, 5).value,
                    worksheet.cell(row, 6).value,
                    worksheet.cell(row, 7).value,
                    worksheet.cell(row, 8).value,
                    worksheet.cell(row, 9).value,
                    worksheet.cell(row, 10).value,
                    worksheet.cell(row, 11).value,
                    worksheet.cell(row, 12).value,
                    worksheet.cell(row, 13).value,
                    worksheet.cell(row, 14).value,
                    worksheet.cell(row, 15).value,
                    worksheet.cell(row, 16).value,
                ))


def accum_parse(region, date, worksheet, rtus):
    rtu_dict = {}
    for rtu in rtus:
        rtu_dict[rtu] = []

    for i, entry in enumerate(worksheet.col(1)):
        if i:
            try:
                rtu_dict[entry.value].append(i)
            except:
                pass

    for i, rtu in enumerate(rtus):
        print('accumulator {}/{}'.format(i + 1, len(rtus)))
        outfile_dir = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + rtu + '\\'
        if not os.path.exists(outfile_dir):
            os.makedirs(outfile_dir)

        outfile_name = date + '_' + rtu + '_ACCUM.csv'
        outfile = outfile_dir + outfile_name

        with open(outfile, 'w+') as output_file:
            output_file.write('STATION,RTU,TYPE_RTU,RTU ACCUMULATOR,PHYADR_PULSE,EMS ACCUMULATOR,PRI SITE,SEC SITE2,SCALE_PULSE,ID_DEVICE (short),NAME_DEVICE (descriptive)\n')
            for row in rtu_dict[rtu]:
                output_file.write('{},{},{},{},{},{},{},{},{},{},{}\n'.format(
                    worksheet.cell(row, 0).value,
                    worksheet.cell(row, 1).value,
                    worksheet.cell(row, 2).value,
                    worksheet.cell(row, 3).value,
                    worksheet.cell(row, 4).value,
                    worksheet.cell(row, 5).value,
                    worksheet.cell(row, 6).value,
                    worksheet.cell(row, 7).value,
                    worksheet.cell(row, 8).value,
                    worksheet.cell(row, 9).value,
                    worksheet.cell(row, 10).value,
                ))


def anout_parse(region, date, worksheet, rtus):
    rtu_dict = {}
    for rtu in rtus:
        rtu_dict[rtu] = []

    for i, entry in enumerate(worksheet.col(1)):
        if i:
            try:
                rtu_dict[entry.value].append(i)
            except:
                pass

    for i, rtu in enumerate(rtus):
        print('analog out {}/{}'.format(i + 1, len(rtus)))
        outfile_dir = '\\\\bmcd\\dfs\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\' + region + '\\' + rtu + '\\'
        if not os.path.exists(outfile_dir):
            os.makedirs(outfile_dir)

        outfile_name = date + '_' + rtu + '_ANOUT.csv'
        outfile = outfile_dir + outfile_name

        with open(outfile, 'w+') as output_file:
            output_file.write('STATION,RTU,TYPE_RTU,RTU ACCUMULATOR,PHYADR_PULSE,EMS ACCUMULATOR,PRI SITE,SEC SITE2,SCALE_PULSE,ID_DEVICE (short),NAME_DEVICE (descriptive)\n')
            for row in rtu_dict[rtu]:
                output_file.write('{},{},{},{},{},{},{},{}\n'.format(
                    worksheet.cell(row, 0).value,
                    worksheet.cell(row, 1).value,
                    worksheet.cell(row, 2).value,
                    worksheet.cell(row, 3).value,
                    worksheet.cell(row, 4).value,
                    worksheet.cell(row, 5).value,
                    worksheet.cell(row, 6).value,
                    worksheet.cell(row, 7).value,
                ))


def ems_parse(region, date, workbook, rtus):
    archive_fep_files(region, rtus)
    archive_rtu_files(region, workbook.sheet_by_name('BMCD_RTUC and RTU'), rtus)
    rtu_fep_parse(region, date, workbook.sheet_by_name('BMCD_RTUC and RTU'), rtus)
    status_parse(region, date, workbook.sheet_by_name('BMCD_STATUS'), rtus)
    control_parse(region, date, workbook.sheet_by_name('BMCD_CONTROL'), rtus)
    analog_parse(region, date, workbook.sheet_by_name('BMCD_ANALOG'), rtus)
    accum_parse(region, date, workbook.sheet_by_name('BMCD_ACCUM'), rtus)
    anout_parse(region, date, workbook.sheet_by_name('BMCD_ANOUT'), rtus)


if __name__ == '__main__':
    s_time = time.time()
    file_dir = Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    file_full_path = askopenfilename(title='Select EMS Dump Excel file')  # show an "Open" dialog box and return the path to the selected file
    # file_full_path = 'Z:/Clients/TND/FirstEnr/82568_EtfScadaSupprt/Design/Substation Projects/EMS MODEL SCREEN DUMPS/20180301 - SOUTH - SNAPSHOT - TELEMETRY CROSS-REF'
    # file_full_path = 'C:/Users/machristiansen/Desktop/20180301 - SOUTH - SNAPSHOT - TELEMETRY CROSS-REF.xlsx'
    # file_full_path = 'C:/Users/cldavis3/Desktop/20180830 - EAST - SNAPSHOT - TELEMETRY CROSS-REF.xlsx'

    # find where excel file name starts and grab the file name + date of EMS upload dump
    fname_index = file_full_path.rfind('/')
    filename = file_full_path[fname_index+1:]
    dump_date = filename[:8]

    # decide if east, west, or south
    if filename[11:12] == 'E':
        ems_region = 'EAST'
    elif filename[11:12] == 'W':
        ems_region = 'WEST'
    else:
        ems_region = 'SOUTH'

    # load EMS dump excel workbook and create List of RTU names
    print(ems_region)
    print('Opening workbook...')
    wbook = xlrd.open_workbook(file_full_path)
    wsheet = wbook.sheet_by_index(0)
    print('Grabbing RTU list...')
    rtu_list = grab_rtu_list(wsheet)

    # parse through dump file for each RTU
    print('Beginning parse of spreadsheet...')
    ems_parse(ems_region, dump_date, wbook, rtu_list)
    f_time = time.time()
    print(f_time-s_time)
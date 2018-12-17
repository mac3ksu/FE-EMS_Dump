import os
import shutil
import time
import xlrd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import zipfile

def grab_rtu_list(worksheet):
    rtus_raw = worksheet.col(2)
    rtus = []
    for rtu in rtus_raw:
        if rtu.value != '' and rtu.value != '#N/A':
            rtus.append(str(rtu.value))
    rtus.pop(0)
    rtus.sort()
    print(rtus)
    return rtus


def create_site_folders(region, rtus, directory):
    outfile_dir = os.path.join(directory, region, '_RTU_FEP')
    if not os.path.exists(outfile_dir):
        os.makedirs(outfile_dir)
    for i, rtu in enumerate(rtus):
        print('creating {} folder {}/{}'.format(rtu, i+1, len(rtus)))
        outfile_dir = os.path.join(directory, region, rtu)
        if not os.path.exists(outfile_dir):
            os.makedirs(outfile_dir)


def archive_rtu_files(region, rtus, directory):
    # for all rtus in the ems dump excel doc, archive existing csv files
    for i, rtu in enumerate(rtus):
        print('archive {}/{}'.format(i + 1, len(rtus)))
        site_dir = os.path.join(directory, region, rtu)
        # Creates an empty archive zip file if no zip file exists
        with zipfile.ZipFile(os.path.join(site_dir, region + '_' + rtu + '_ARCHIVE' + '.zip'), 'a', zipfile.ZIP_DEFLATED) as myzip:
            try:
                # Iterates through files in site directory, if the file is a CSV file extension it is appended
                # to the archive zip and then deleted
                for file in os.listdir(site_dir):
                    if str(file[-3:]) == 'csv':
                        myzip.write(os.path.join(site_dir, file), file)
                        os.remove(os.path.join(site_dir, file))
            except:
                pass
        myzip.close()


def rtu_fep_parse(region, date, worksheet, directory):
    outfile_dir = os.path.join(directory, region, '_RTU_FEP')
    outfile_name = date + '_' + region + '_RTU_FEP.csv'
    outfile = os.path.join(outfile_dir, outfile_name)

    i = 0

    with open(outfile, 'w+') as output_file:
        while i < worksheet.nrows:
            if worksheet.cell(i,0).value == '':
                break
            else:
                output_file.write('{},{},{},{},{},{},{},{},{},{},{},{},{},{}\n'.format(
                    worksheet.cell(i, 0).value,
                    worksheet.cell(i, 1).value,
                    worksheet.cell(i, 2).value,
                    worksheet.cell(i, 3).value,
                    worksheet.cell(i, 4).value,
                    worksheet.cell(i, 5).value,
                    worksheet.cell(i, 6).value,
                    worksheet.cell(i, 7).value,
                    worksheet.cell(i, 8).value,
                    worksheet.cell(i, 9).value,
                    worksheet.cell(i, 10).value,
                    worksheet.cell(i, 11).value,
                    worksheet.cell(i, 12).value,
                    worksheet.cell(i, 13).value,
                ))
            i += 1



def status_parse(region, date, worksheet, rtus, directory):
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
        if len(rtu_dict[rtu]) > 0:
            print('status {} {}/{}'.format(rtu, i+1, len(rtus)))
            # print(rtu)
            outfile_dir = directory + '\\' + region + '\\' + rtu + '\\'

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


def control_parse(region, date, worksheet, rtus, directory):
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
        if len(rtu_dict[rtu]) > 0:
            print('control {} {}/{}'.format(rtu, i + 1, len(rtus)))
            outfile_dir = directory + '\\' + region + '\\' + rtu + '\\'

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


def analog_parse(region, date, worksheet, rtus, directory):
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
        if len(rtu_dict[rtu]) > 0:
            print('analog {} {}/{}'.format(rtu, i + 1, len(rtus)))
            outfile_dir = directory + '\\' + region + '\\' + rtu + '\\'

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


def accum_parse(region, date, worksheet, rtus, directory):
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
        if len(rtu_dict[rtu])>0:
            print('accumulator {} {}/{}'.format(rtu, i + 1, len(rtus)))
            outfile_dir = directory + '\\' + region + '\\' + rtu + '\\'

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


def anout_parse(region, date, worksheet, rtus, directory):
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
        if len(rtu_dict[rtu]) > 0:
            print('analog out {} {}/{}'.format(rtu, i + 1, len(rtus)))
            outfile_dir = directory + '\\' + region + '\\' + rtu + '\\'

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


def ems_parse(region, date, workbook, rtus, directory):
    create_site_folders(region, rtus, directory)
    archive_rtu_files(region, rtus, directory)
    rtu_fep_parse(region, date, workbook.sheet_by_name('BMCD_RTUC and RTU'), directory)
    status_parse(region, date, workbook.sheet_by_name('BMCD_STATUS'), rtus, directory)
    control_parse(region, date, workbook.sheet_by_name('BMCD_CONTROL'), rtus, directory)
    analog_parse(region, date, workbook.sheet_by_name('BMCD_ANALOG'), rtus, directory)
    accum_parse(region, date, workbook.sheet_by_name('BMCD_ACCUM'), rtus, directory)
    anout_parse(region, date, workbook.sheet_by_name('BMCD_ANOUT'), rtus, directory)


if __name__ == '__main__':
    s_time = time.time()
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    file_full_path = askopenfilename(
        title='Select EMS Dump Excel file')  # show an "Open" dialog box and return the path to the selected file

    # find where excel file name starts and grab the file name + date of EMS upload dump
    fname_index = file_full_path.rfind('/')
    filename = file_full_path[fname_index+1:]
    print(filename)
    file_dir = os.path.normpath(file_full_path[:fname_index])
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
    ems_parse(ems_region, dump_date, wbook, rtu_list, file_dir)
    f_time = time.time()
    print(str(int(f_time-s_time)/60) + ' minutes')

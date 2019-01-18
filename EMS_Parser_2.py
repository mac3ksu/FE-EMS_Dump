import os
import shutil
import xlrd
import time
#

s_time = time.time()

root_src_dir = 'C:\\Users\\machristiansen\\Desktop\\EMS Dump Test\\WEST'
root_dst_dir = 'Z:\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\WEST'

for dst_dir, dirs, files in os.walk(root_dst_dir):
    for file_ in files:
        if str(file_[-3:]) == 'csv':
            print('removing WEST ' + file_)
            os.remove(os.path.join(dst_dir, file_))



for src_dir, dirs, files in os.walk(root_src_dir):
    dst_dir = src_dir.replace(root_src_dir, root_dst_dir, 1)
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    for file_ in files:
        src_file = os.path.join(src_dir, file_)
        dst_file = os.path.join(dst_dir, file_)
        print(dst_file)
        if os.path.exists(dst_file):
            os.remove(dst_file)
        shutil.copy(src_file, dst_dir)

root_src_dir = 'C:\\Users\\machristiansen\\Desktop\\EMS Dump Test\\EAST'
root_dst_dir = 'Z:\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\EAST'

for dst_dir, dirs, files in os.walk(root_dst_dir):
    for file_ in files:
        if str(file_[-3:]) == 'csv':
            print('removing EAST ' + file_)
            os.remove(os.path.join(dst_dir, file_))



for src_dir, dirs, files in os.walk(root_src_dir):
    dst_dir = src_dir.replace(root_src_dir, root_dst_dir, 1)
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    for file_ in files:
        src_file = os.path.join(src_dir, file_)
        dst_file = os.path.join(dst_dir, file_)
        print(dst_file)
        if os.path.exists(dst_file):
            os.remove(dst_file)
        shutil.copy(src_file, dst_dir)

root_src_dir = 'C:\\Users\\machristiansen\\Desktop\\EMS Dump Test\\SOUTH'
root_dst_dir = 'Z:\\Clients\\TND\\FirstEnr\\82568_EtfScadaSupprt\\Design\\Substation Projects\\EMS MODEL SCREEN DUMPS\\SOUTH'

for dst_dir, dirs, files in os.walk(root_dst_dir):
    for file_ in files:
        if str(file_[-3:]) == 'csv':
            print('removing SOUTH ' + file_)
            os.remove(os.path.join(dst_dir, file_))



for src_dir, dirs, files in os.walk(root_src_dir):
    dst_dir = src_dir.replace(root_src_dir, root_dst_dir, 1)
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    for file_ in files:
        src_file = os.path.join(src_dir, file_)
        dst_file = os.path.join(dst_dir, file_)
        print(dst_file)
        if os.path.exists(dst_file):
            os.remove(dst_file)
        shutil.copy(src_file, dst_dir)

f_time = time.time()
print(str(int(f_time-s_time)/60) + ' minutes')

# wbook = xlrd.open_workbook('C:\\Users\\machristiansen\\Desktop\\EMS Dump Test\\20180801 - SOUTH - SNAPSHOT - TELEMETRY CROSS-REF.xlsx')
# wsheet = wbook.sheet_by_name('BMCD_RTUC and RTU')
# print('num rows ' + str(wsheet.nrows))
# print('num cols ' + str(wsheet.ncols))

from tkinter import filedialog
from tkinter import *
import time
import xlrd
import os
import datetime
from tkinter import messagebox
import tkinter.font as font

cwd = os.getcwd()

def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    global ciqfiles
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    ciqfiles = os.listdir(folder_path.get())

def complete():
    messagebox.showinfo(title='Inventory', message='Inventory files generated successfully!', )

def inventory():
    global batch
    batch1 = batch.get()
    ciqcount = len(ciqfiles)
    cwd = os.getcwd()
    today = str(datetime.date.today())

    region_dic = {'northeast': 'ne', 'southeast': 'se', 'southwest': 'sw', 'central': 'c', 'west1': 'w1', 'west2': 'w2'}
    if os.path.exists(cwd + '\\batch' + batch1):
        pass
    else:
        os.mkdir(cwd + '\\batch' + batch1)
    if os.path.exists(cwd + '\\batch' + batch1 + '\\' + today):
        filelist = [f for f in os.listdir(cwd + '\\batch' + batch1 + '\\' + today)]
        for f in filelist:
            os.remove(cwd + '\\batch' + batch1 + '\\' + today + '/' + f)
        print('Old files cleaned up!')
    else:
        os.mkdir(cwd + '\\batch' + batch1 + '\\' + today)

    i = 0
    while i < ciqcount:
        workbook = xlrd.open_workbook(folder_path.get() + '\\' + ciqfiles[i])
        worksheet = workbook.sheet_by_name('IP')
        podname = worksheet.cell(1, 2)
        region = worksheet.cell(3, 2)
        hostnamefile = open(cwd + '\\batch' + batch1 + '\\' + today + '\\' + podname.value, 'a')
        hostnamefile.write('localhost ansible_connection=local\n\n[pod]\n')
        hostnamefile.write(str(podname.value).lower())
        hostnamefile.write('\n\n[mgmt]\n')
        ctswitch = worksheet.cell(24, 1)
        hostnamefile.write(str(ctswitch.value).rstrip('1').lower() + '[1:2]\n\n[spines]\n')

        spswitch = worksheet.cell(26, 1)
        hostnamefile.write(str(spswitch.value).rstrip('1').lower() + '[1:2]\n\n[leaves]\n')

        lfswitch = worksheet.cell(28, 1)
        hostnamefile.write(str(lfswitch.value).rstrip(
            '1').lower() + '[1:4]\n\n[nxos:children]\nspines\nleaves\n\n[switches:children]\nmgmt\nspines\nleaves\n\n[ospd]\n')

        ospdsrvr = worksheet.cell(33, 1)
        ospdsrvr = ospdsrvr.value
        ospdsrvr1 = str(ospdsrvr).replace('UCS', '')
        ospdsrvr2 = str(ospdsrvr1).replace('-', '')
        ospdsrvr3 = str(ospdsrvr2).replace('(RHEL7)', '')
        hostnamefile.write(ospdsrvr3.lower() + '\n\n[saegw_c]\n')

        saegw_clist = []
        for saegw_c in range(49, 51):
            vnfssaegw_c = worksheet.cell(saegw_c, 1)
            saegw_clist.append(vnfssaegw_c.value)

        pgw_clist = []
        for pgw_c in range(53, 55):
            vnfspgw_c = worksheet.cell(pgw_c, 1)
            pgw_clist.append(vnfspgw_c.value)

        saegw_chalist = []
        for saegw_cha in range(49, 51):
            vnfssaegw_cha = worksheet.cell(saegw_cha, 2)
            saegw_chalist.append(vnfssaegw_cha.value)

        pgw_chalist = []
        for pgw_cha in range(53, 55):
            vnfspgw_cha = worksheet.cell(pgw_cha, 2)
            pgw_chalist.append(vnfspgw_cha.value)

        saegw_clistlen = len(saegw_clist)
        saegwc = 0
        while saegwc < saegw_clistlen:
            hostnamefile.write(
                str(saegw_clist[saegwc]).lower() + ' ha=' + str(saegw_chalist[saegwc]).split('-')[0].lower() + '\n')
            saegwc = saegwc + 1

        hostnamefile.write('\n[pgw_c]\n')

        pgw_clistlen = len(pgw_clist)
        pgwc = 0
        while pgwc < pgw_clistlen:
            hostnamefile.write(
                str(pgw_clist[pgwc]).lower() + ' ha=' + str(pgw_chalist[pgwc]).split('-')[0].lower() + '\n')
            pgwc = pgwc + 1

        hostnamefile.write('\n[saegw_u]\n')

        saegw_ulist = []
        for saegw_u in range(55, 59):
            vnfssaegw_u = worksheet.cell(saegw_u, 1)
            saegw_ulist.append(vnfssaegw_u.value)

        sgw_ulist = []
        for sgw_u in range(59, 65):
            vnfssgw_u = worksheet.cell(sgw_u, 1)
            sgw_ulist.append(vnfssgw_u.value)

        pgw_ulist = []
        for pgw_u in range(65, 73):
            vnfspgw_u = worksheet.cell(pgw_u, 1)
            pgw_ulist.append(vnfspgw_u.value)

        upf_pgw_ulist = []
        for upf_pgw_u in range(73, 87):
            vnfsupf_pgw_u = worksheet.cell(upf_pgw_u, 1)
            upf_pgw_ulist.append(vnfsupf_pgw_u.value)


        saegw_uhalist = []
        for saegw_uha in range(55, 59):
            vnfssaegw_uha = worksheet.cell(saegw_uha, 2)
            saegw_uhalist.append(vnfssaegw_uha.value)

        sgw_uhalist = []
        for sgw_uha in range(59, 65):
            vnfssgw_uha = worksheet.cell(sgw_uha, 2)
            sgw_uhalist.append(vnfssgw_uha.value)
        pgw_uhalist = []
        for pgw_uha in range(65, 73):
            vnfspgw_uha = worksheet.cell(pgw_uha, 2)
            pgw_uhalist.append(vnfspgw_uha.value)
        upf_pgw_uhalist = []
        for upf_pgw_uha in range(73, 87):
            vnfsupf_pgw_uha = worksheet.cell(upf_pgw_uha, 2)
            upf_pgw_uhalist.append(vnfsupf_pgw_uha.value)

        saegw_ulistlen = len(saegw_ulist)
        saegwu = 0
        while saegwu < saegw_ulistlen:
            hostnamefile.write(
                str(saegw_ulist[saegwu]).lower() + ' ha=' + str(saegw_uhalist[saegwu]).split('-')[0].lower() + '\n')
            saegwu = saegwu + 1
        hostnamefile.write('\n[sgw_u]\n')
        sgw_ulistlen = len(sgw_ulist)
        sgwu = 0
        while sgwu < sgw_ulistlen:
            hostnamefile.write(
                str(sgw_ulist[sgwu]).lower() + ' ha=' + str(sgw_uhalist[sgwu]).split('-')[0].lower() + '\n')
            sgwu = sgwu + 1
        hostnamefile.write('\n[pgw_u]\n')
        pgw_ulistlen = len(pgw_ulist)
        pgwu = 0
        while pgwu < pgw_ulistlen:
            hostnamefile.write(
                str(pgw_ulist[pgwu]).lower() + ' ha=' + str(pgw_uhalist[pgwu]).split('-')[0].lower() + '\n')
            pgwu = pgwu + 1
        hostnamefile.write('\n[upf_pgw_u]\n')

        upf_pgw_ulistlen = len(upf_pgw_ulist)
        upfpgwu = 0
        while upfpgwu < upf_pgw_ulistlen:
            hostnamefile.write(str(upf_pgw_ulist[upfpgwu]).lower() + ' ha=' + str(upf_pgw_uhalist[upfpgwu]).split('-')[
                0].lower() + '\n')
            upfpgwu = upfpgwu + 1

        hostnamefile.write('''\n[vnfs:children]
saegw_c
saegw_u
pgw_c
pgw_u
sgw_u
upf_pgw_u\n\n[smf_ims]\n''')

        smf_ims = worksheet.cell(91, 1)
        smf_ims_ip = worksheet.cell(91, 5)
        hostnamefile.write(
            str(smf_ims.value).lower() + ' ansible_host=' + str(smf_ims_ip.value).lower() + '\n\n[smf_data]\n')

        smf_data = worksheet.cell(99, 1)
        smf_data_ip = worksheet.cell(99, 5)
        hostnamefile.write(
            str(smf_data.value).lower() + ' ansible_host=' + str(smf_data_ip.value).lower() + '\n\n[master_vip]\n')

        mastervip = worksheet.cell(111, 1)
        mastervip_ip = worksheet.cell(111, 5)
        hostnamefile.write(str(mastervip.value).lower() + ' ansible_host=' + mastervip_ip.value + '\n\n')

        hostnamefile.write('''[vms:children]
smf_ims
smf_data
master_vip\n\n''')

        podshortname = worksheet.cell(2, 2)

        hostnamefile.write('[' + str(podshortname.value).lower() + ']\n')
        hostnamefile.write(str(ctswitch.value).rstrip('1').lower() + '[1:2]\n')
        hostnamefile.write(str(spswitch.value).rstrip('1').lower() + '[1:2]\n')
        hostnamefile.write(str(lfswitch.value).rstrip('1').lower() + '[1:4]\n')
        hostnamefile.write(ospdsrvr3.lower() + '\n')
        saegw_c_beg = str(saegw_clist[0])[6:]
        saegw_c_end = str(saegw_clist[-1])[6:]
        hostnamefile.write(str(podshortname.value)[0:2].lower() + 'pcf5[' + saegw_c_beg + ':' + saegw_c_end + ']\n')

        pgw_c_beg = str(pgw_clist[0])[6:]
        pgw_c_end = str(pgw_clist[-1])[6:]
        hostnamefile.write(str(podshortname.value)[0:2].lower() + 'pgw5[' + pgw_c_beg + ':' + pgw_c_end + ']\n')

        saegw_ubeg = str(saegw_ulist[0])[6:]
        saegw_uend = str(saegw_ulist[-1])[6:]
        hostnamefile.write(str(podshortname.value)[0:2].lower() + 'pcf6[' + saegw_ubeg + ':' + saegw_uend + ']\n')

        sgw_u_beg = str(sgw_ulist[0])[6:]
        sgw_u_end = str(sgw_ulist[-1])[6:]
        hostnamefile.write(str(podshortname.value)[0:2].lower() + 'sgw6[' + sgw_u_beg + ':' + sgw_u_end + ']\n')

        pgw_u_beg = str(pgw_ulist[0])[6:]
        pgw_u_end = str(pgw_ulist[-1])[6:]
        hostnamefile.write(str(podshortname.value)[0:2].lower() + 'pgw6[' + pgw_u_beg + ':' + pgw_u_end + ']\n')

        upf_pgw_u_beg = str(upf_pgw_ulist[0])[6:]
        upf_pgw_u_end = str(upf_pgw_ulist[-1])[6:]
        hostnamefile.write(str(podshortname.value)[0:2].lower() + 'upf0[' + upf_pgw_u_beg + ':' + upf_pgw_u_end + ']\n')

        hostnamefile.write(str(smf_ims.value).lower() + '\n')
        hostnamefile.write(str(smf_data.value).lower() + '\n')
        hostnamefile.write(str(mastervip.value).lower() + '\n')
        hostnamefile.write(str(podname.value).lower() + '\n\n')
        hostnamefile.write('[' + str(podshortname.value).lower() + ':vars]\n')
        hostnamefile.write('pod_id=' + str(podshortname.value)[0:2].lower() + 'ucs' + str(podshortname.value)[2:] + '\nshort_pod_id=' + str(podshortname.value).lower() + '\nbatch_id=' + batch1 + '\nregion=' + region_dic[str(region.value).lower()] + '\n')

        n7_vip = worksheet.cell(328, 7)
        n40_vip = worksheet.cell(330, 7)
        nrf_dns_vip = worksheet.cell(332, 3)

        hostnamefile.write('n7_vip=' + n7_vip.value + '\n')
        hostnamefile.write('n40_vip=' + n40_vip.value + '\n')
        hostnamefile.write('nrf_dns_vip=' + nrf_dns_vip.value + '\n')

        i = i + 1
    time.sleep(2)
    hostnamefile.close()

    time.sleep(2)
    complete()


root = Tk()
myFont = font.Font(family='Helvetica', size=11, weight='bold')
root.title('Ansible Inventory')
root.geometry('400x200')
folder_path = StringVar()
lbl1 = Label(master=root,textvariable=folder_path)
lbl1.place(x=200, y=25)
button2 = Button(text="CIQ Folder", width=12, font=myFont, bg='#B30095', fg='#ffffff', command=browse_button)
button2.place(x=15, y=25)

button3 = Button(text="Inventory", width=12, font=myFont, bg='#125eee', fg='#ffffff', command=inventory)
button3.place(x=15, y=75)

batch = Entry(master=root, font=myFont)
batch.place(x=15, y=125)

mainloop()

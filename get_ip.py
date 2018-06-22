#！ -*- coding:utf-8 -*-

import os
import re
import sys
import xlrd
from xlutils.copy import copy


def get_ip_by_domain(domain):
    """
    根据domain获取ip
    """
    #判断是否已经是一个ip，而不是domain
    if re.match(r"^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$", domain):
        return domain    

    mcmd = "nslookup %s" % domain
    mp = os.popen(mcmd)
    res = mp.read()

    tmp = res.split(": ")
    tmp_list = tmp[4].split("\n")
     
    res_list = []
    for item in tmp_list:
        if item != "" and item != "Aliases":
            res_list.append(item.strip(" \t"))

    return "; ".join(res_list)


def load_excel(filename, scol):
    wbook = xlrd.open_workbook(filename)
    wbsheet = wbook.sheet_by_index(0)
    nrow = wbsheet.nrows

    domain_list = []
    for idx in range(1, nrow):
        value = wbsheet.cell(idx, scol).value
        domain_list.append(value)

    return domain_list


def update_excel(filename, scol, ip_list):
    oldwb = xlrd.open_workbook(filename)
    wb = copy(oldwb)
    ws = wb.get_sheet(0)

    for idx in range(0, len(ip_list)):
        ws.write(idx+1, scol, ip_list[idx])

    wb.save("./result.xlsx")

def modify_entrance(filename, scol=2):
    if not os.path.exists(filename):
        print("%s not exists" % filename)
        return -1
    domain_list = load_excel(filename, scol)

    ip_list = []
    for item in domain_list:
        ip = get_ip_by_domain(item.split("//")[1].strip())
        print("%s : %s" % (item , ip))
        ip_list.append(ip)
    
    update_excel(filename, scol + 1, ip_list)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: python get_ip.py [d|f] domain|filename")
    else:
        var = sys.argv[1]
        if var == "d":
            domain = sys.argv[2]
            res = get_ip_by_domain(domain)
            print(res)
        elif var == "f":
            filename = sys.argv[2]
            modify_entrance(filename)
#PDF Report Crawler V1.0#
# @ycooi ooiyungchaw@hotmail.com #


import os
import re
import types
import time
import xlwt
EXE_PATH = 'J:\Total_PDF_Converter\PDFConverter.exe'


XLS_ROWS = []
temp = ['Type', 'Prilled Urea - fob bulk', 'Prilled Urea - fob bulk', 'Prilled Urea - fob bulk', 'Prilled Urea - fob bulk', 'Prilled Urea - fob bulk', 'Prilled Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Granular Urea - fob bulk', 'Ammonium Sulphate - bulk', 'Ammonium Sulphate - bulk', 'Ammonium Sulphate - bulk', 'Ammonium Sulphate - bulk', 'Ammonium Sulphate - bulk', 'Ammonium Sulphate - bulk', 'Ammonium Nitrate', 'Ammonium Nitrate', 'Ammonium Nitrate', 'Ammonium Nitrate', 'Ammonium Nitrate']
XLS_ROWS.append(temp)

temp = ['Date', 'Black Sea', 'Baltic', 'Arabian Gulf', 'China', 'Croatia/Romania', 'Brazil (cfr)', 'Arabian Gulf all netbacks', 'Arabian Gulf US netback', 'Arabian Gulf non-US netback', 'Iran', 'Egypt', 'Algeria', 'Noth Africa full range', 'China', 'Indonesia/Malaysia', 'Southeast Asia (cfr)', 'Venezuela/Trinidad', 'Brazil (cfr)', 'US Gulf (pst barge)', 'US Gulf (cfr metric)', 'French Atlantic (fca Euro)', 'Baltic', 'fob Baltic (caprolactam)', 'fob Blk Sea (caprolactam)', 'fob Kherson (steel grade)', 'fob China (caprolactam)', 'cfr S.E. Asia (caprolactam)', 'cfr Brazil (caprolactam)', 'fob bulk Baltic', 'fob bulk Black Sea', 'France (fca euros bagged)', 'UK (fca sterling bagged)', 'CAN 27 Germany (cif inland euros)']
XLS_ROWS.append(temp)


PDF_FILE = 'a.pdf'
PDF_FILE_DIR = 'pdf/'

NUM_PAT = re.compile('\d+', re.I)

SETTING = []
item_type = ['Prilled Urea fob bulk']
items = ['black sea', 'baltic', 'Arabian Gulf', 'China', 'Croatia Romania', 'Brazil cfr']
SETTING.append({'type':item_type, 'items':items})


item_type = ['Granular Urea fob bulk']
items = ['Arabian Gulf all netbacks', 'Arabian Gulf US netback', 'Arab Gulf non US netbacks', 'Iran', 'Egypt', 'Algeria', 'Noth Africa full range', 'China', 'Indonesia Malaysia', 'Southeast Asia cfr', 'Venezuela Trinidad', 'Brazil cfr', 'US Gulf p.s.t. barge', 'US Gulf cfr metric', 'French Atlantic fca Euro', 'Baltic']
SETTING.append({'type':item_type, 'items':items})

item_type = ['Ammonium Sulphate bulk']
# items = ['fob Baltic caprolactam', 'fob Blk Sea caprolactam', 'fob Kherson steel grade', 'fob China caprolactam', 'cfr S.E. Asia caprolactam', 'cfr Brazil caprolactam']
items = ['fob Baltic caprolactam', 'fob Blk Sea caprolactam', 'fob Kherson steel grade', 'fob China caprolactam', 'cfr S.E. Asia', 'cfr Brazil caprolactam']
SETTING.append({'type':item_type, 'items':items})

item_type = ['Ammonium Nitrate']
items = ['fob bulk Baltic', 'fob bulk Black Sea', 'France fca euros bagged', 'UK fca sterling bagged', 'CAN 27 Germany cif inland euros']
SETTING.append({'type':item_type, 'items':items})


PRICE_PAT = re.compile(r'[0-9\-]+', re.IGNORECASE)

WRITE_CNT = 0
COL_CNT = 0

#############################################################################################
SETTING2015 = []
item_type = ['Prilled Urea fob bulk']
items = ['black sea', 'baltic', 'Arabian Gulf', 'China', 'Croatia Romania', 'Brazil cfr']
SETTING2015.append({'type':item_type, 'items':items})

item_type = ['Granular Urea fob bulk']
items = ['Arabian Gulf all netbacks', 'Arabian Gulf US netback', 'Arab Gulf non US netbacks', 'Iran', 'Egypt', 'Algeria', 'Noth Africa full range', 'China', 'Indonesia Malaysia', 'Southeast Asia cfr', 'Venezuela Trinidad', 'Brazil cfr', 'US Gulf p.s.t. barge', 'US Gulf cfr metric', 'French Atlantic fca Euro', 'Baltic']
SETTING2015.append({'type':item_type, 'items':items})

item_type = ['Ammonium Sulphate bulk']
items = ['fob Baltic caprolactam', 'fob Blk Sea caprolactam', 'fob Kherson steel grade', 'fob China caprolactam', 'cfr S.E. Asia', 'cfr Brazil caprolactam']
SETTING2015.append({'type':item_type, 'items':items})

item_type = ['Ammonium Nitrate']
items = ['fob bulk Baltic', 'fob bulk Black Sea', 'France fca euros bagged', 'UK fca sterling bagged', 'CAN 27 Germany cif inland euros']
SETTING2015.append({'type':item_type, 'items':items})

item_type = ['UAN 32%']
items = ['nola short ton', 'rouen 30% N fot euro', 'fob black sea', 'fob baltic']
SETTING2015.append({'type':item_type, 'items':items})

##############################################################################################
SETTING2015B = []
item_type = ['DAP MAP TSP fob bulk']
items = ['DAP Tampa', 'DAP Tunisia', 'DAP Morocco', 'DAP Baltic Black Sea', 'DAP China', 'DAP Saudi Arabia KSA', 'DAP Mexico', 'DAP Australia', 'DAP US Gulf domestic barge', 'DAP Central Florida railcar', 'DAP China ex-works', 'DAP Benelux fot fob duty paid free', 'MAP Baltic', 'MAP Morocco', 'TSP Tunisia', 'TSP Morocco', 'TSP China', 'TSP eastern Med Lebanon Israel']
SETTING2015B.append({'type':item_type, 'items':items})

item_type = ['DAP MAP cfr bulk']
items = ['DAP MAP Argentina Uruguay', 'MAP Brazil', 'DAP India contract', 'DAP Pakistan']
SETTING2015B.append({'type':item_type, 'items':items})

item_type = ['NPK 16-16-16 bulk']
items = ['fob FSU', 'cfr China', 'cfr southeast Asia']
SETTING2015B.append({'type':item_type, 'items':items})

item_type = ['Phosphoric acid t P2O5']
items = ['cfr India', 'cfr western Europe', 'cfr Brazil']
SETTING2015B.append({'type':item_type, 'items':items})

item_type = ['Phosphate rock BPL']
items = ['fob Jordan 68 70', 'cfr India 68 70', 'cfr India 70 72']
SETTING2015B.append({'type':item_type, 'items':items})

item_type = ['Sulphur']
items = ['cfr Tampa', 'cfr north Africa']
SETTING2015B.append({'type':item_type, 'items':items})

item_type = ['Ammonia']
items = ['cfr Tampa']
SETTING2015B.append({'type':item_type, 'items':items})
##############################################################################################
SETTING2015C = []
item_type = ['Ammonia fob']
items = ['Ventspils', 'Yuzhny', 'North Africa', 'Middle East', 'US Gulf domestic barge st', 'Caribbean']
SETTING2015C.append({'type':item_type, 'items':items})

item_type = ['Ammonia cfr']
items = ['NW Eur duty unpaid', 'NW Eur duty paid free', 'North Africa', 'India', 'East Asia excl Taiwan', 'Taiwan', 'Tampa', 'Gulf']
SETTING2015C.append({'type':item_type, 'items':items})

##############################################################################################
SETTING2015D = []
item_type = ['Sulphur dry bulk']
items = ['fob Vancouver Q4-2015', 'fob Middle East Q4-2015', 'fob Qatar QSP Dec 2015', 'fob UAE OSP Dec 2015', 'fob Iran', 'fob Black Sea lump gran Q4-2015', 'fob US Gulf Q4-2015', 'cfr Brazil Q4-2015', 'cfr Med under 10 k', 'fob Med under 10 k', 'cfr N Africa lump gran Q4-2015', 'cfr India', 'cfr China Q4-2015', 'ex w Nantong CNY']
SETTING2015D.append({'type':item_type, 'items':items})

item_type = ['Sulphur molten']
items = ['cfr Tampa C Fla Q4-2015', 'cfr Benelux loc refs Q4-2015', 'cpt NW Europe Q4-2015']
SETTING2015D.append({'type':item_type, 'items':items})
##############################################################################################
SETTING2015E = []
item_type = ['MOP fob standard bulk']
items = ['Vancouver fob', 'NW Europe fob', 'FSU fob', 'Jordan fob', 'Israel fob', 'S E Asia cfr', 'India cfr 180 days']
SETTING2015E.append({'type':item_type, 'items':items})

item_type = ['MOP cfr granular bulk']
items = ['Brazil cfr cash', 'Europe cfr']
SETTING2015E.append({'type':item_type, 'items':items})

item_type = ['SOP fob bulk']
items = ['NW Europe fob']
SETTING2015E.append({'type':item_type, 'items':items})
DES_FILE = 'a.txt'
##############################################################################################

def pdf2txt(pdf='b.pdf'):
    try:
        if os.path.exists(DES_FILE):
            os.remove(DES_FILE)
        command = EXE_PATH + ' ' + pdf + ' ' + DES_FILE + ' -c TXT'
        os.system(command)
    except:
        return False
    return True

def xls_init(xls_handle):
    # write XLS_ROWS to xls
    global WRITE_CNT
    for i in XLS_ROWS:
        for (index, j) in enumerate(i):
            xls_handle.write(WRITE_CNT, index, j)
        WRITE_CNT += 1


def write2xls(xls_handle, final_ret, date=None):
    global WRITE_CNT
    global COL_CNT
    # 把结果依次写到xls中去
    if date:
        xls_handle.write(WRITE_CNT, COL_CNT, date)
        COL_CNT += 1
    for (index, i) in enumerate(final_ret):
        xls_handle.write(WRITE_CNT, COL_CNT, i['value'])
        COL_CNT += 1

def match_detect(pat, text, show=False):
    pat_list = pat.split(' ')
    matched = True
    debug_var = 0
    last_index = 0
    cur_index = 0
    for i in pat_list:
        if i.lower() in text.lower():
            cur_index = text.lower().index(i.lower())
            # if cur_index < last_index:
                # matched = False
                # break
            matched = matched & True
            if False:
                # print last_index, cur_index
                print '--------------------------------------'
                print pat_list
                print i
                print pat
                print text
                debug_var += 1
            last_index = cur_index
            text = text[cur_index:]
        else:
            matched = False
            break
    if debug_var:
        print 'fob' in text
        print pat_list[2] in text
        print pat_list[2]
        print debug_var, matched
    return matched

def line_detail_info(dest, line_text):
    orgin_dest = dest
    dest = dest.lower().split(' ')
    position = line_text.lower().find(dest[0])
    # if dest[0].lower() == 'baltic':
        # print '===================='
        # print dest
        # print line_text
        # print position
        # print line_text[position:]
    price = PRICE_PAT.findall(line_text[position:])
    if not price:
        return ''
    else:
        # for i in price:
            # if i.lower() not in orgin_dest.lower():
                # return i
        # return price[0]
        temp = []
        for i in price:
            if i.lower() not in orgin_dest.lower():
                temp.append(i)
        return temp[0]
        if len(temp) > 2:                                       # 这里注意2015某几个特例
            return temp[0] + '  ' + temp[1]
        else:
            return temp[0]

def remove_bracket(page):
    # pat = re.compile(UNI2UTF(u'（[^（^）]+）'), re.IGNORECASE)
    pat = re.compile('（.+）', re.IGNORECASE)
    ret = pat.findall(page)
    for i in ret:
        b = ''.join([' ' for j in range(len(i))])
        page = page.replace(i, b)

    pat = re.compile('\(.+\)', re.IGNORECASE)
    ret = pat.findall(page)
    for i in ret:
        b = ''.join([' ' for j in range(len(i))])
        page = page.replace(i, b)

    pat = re.compile('（.+\)', re.IGNORECASE)
    ret = pat.findall(page)
    for i in ret:
        b = ''.join([' ' for j in range(len(i))])
        page = page.replace(i, b)

    pat = re.compile('\(.+）', re.IGNORECASE)
    ret = pat.findall(page)
    for i in ret:
        b = ''.join([' ' for j in range(len(i))])
        page = page.replace(i, b)
    return page

def info_extract(dest, text, position=0):
    # dest : items 列表
    # text : 文本列表
    ret_list = []
    dest = dest if type(dest) is types.ListType else [dest]
    for i in dest:
        i = i.lower()                 # 目标 item
        matched = False
        for (index, j) in enumerate(text):
            offset = 0 if position < 40 else 40
            if not match_detect(i, j[offset:], True):       # 加入 offset限定，防止是右边的，确选到左边的了
                continue
            # 还要再做一次检测，防止item其实是另一半的item
            # print i
            # print j
            cur_item = i.lower().split(' ')
            item_index = j[offset:].lower().find(cur_item[0]) + offset
            if item_index == -1:
                continue
            if (item_index - position) > 25:                #  实际是左边，结果找到了右边
                # print '0sddddddddddddddddddddddddddddddd'
                continue
            # if 'brazil' in cur_item:
                # print item_index, position
            # temp_line = ''
            price = line_detail_info(i, j[offset:])
            # print 'price:', price

            # 当前条目得到的结果，条目名称是否完全匹配
            price_temp = price.split(' ')
            act_index = j[offset:].index(price_temp[0]) + offset
            # if act_index < item_index:        # 寻找价格里面已经做了限定，价格必在 item 后面, 这里加上会导致误判
                # continue
            if act_index < (len(j)/2):
                item_name = j[0:act_index].strip()
                item_name = remove_bracket(item_name).strip()
                if (len(item_name) - len(i)) > 6:     # 防止误判, 价格左边的字符必须和item长度符合
                    # print 'item_name:', item_name
                    # print i
                    # print '1sddddddddddddddddddddddddddddddd'
                    continue
            else:
                if j[item_index-2] != ' ':
                    # print '2sddddddddddddddddddddddddddddddd'
                    continue
                # print '------------------'
                # print j[offset:]
                # print j
                # print cur_item[-1]
                item_index = j[offset:].lower().index(cur_item[-1]) + offset    # 如果
                if len(j) > (item_index+len(cur_item[-1])+1):
                    temp = remove_bracket(j)
                    if temp[item_index+len(cur_item[-1])+1] != ' ':
                        # print '3sddddddddddddddddddddddddddddddd'
                        continue
            # 位置判断，防止调到另一半去了
            item_index = j[offset:].lower().index(cur_item[-1]) + offset
            if (act_index - item_index) > 40:
                # print 'WARNING: NO VALUE.', i
                continue
            matched = True
            ret_list.append({'item':i, 'value':price, 'index':index})
            # print ret_list[-1]
            break
        if not matched:
            # print 'NOT MATCHED:', i
            ret_list.append({'item':i, 'value':'', 'index':1})      # 没有找到的，index设为1，防止对均值有干扰
    # 对结果做进一步的修正，防止出现误判
    indexs = [i['index'] for i in ret_list]
    final_ret = []
    for i in ret_list:
        if (i['index'] - min(indexs)) > (30 + len(ret_list)):       # 结果条目离当前表太远了，是别的表下面同名的条目
            # print 'WARNING, TOO FAR:', i
            i['value'] = ''
            # print 'FIXED:', i

        # print i
        final_ret.append(i)
    return final_ret

def xls_init2015(xls_file, init_file):
    global WRITE_CNT
    xls_file.write(WRITE_CNT, 0, 'type')
    col_cnt = 0
    for i in init_file:
        for (index, j) in enumerate(i['items']):
            xls_file.write(WRITE_CNT, col_cnt+1, i['type'])
            col_cnt += 1

    WRITE_CNT += 1
    col_cnt = 0
    xls_file.write(WRITE_CNT, 0, 'date')
    for i in init_file:
        for (index, j) in enumerate(i['items']):
            xls_file.write(WRITE_CNT, col_cnt+1, j)
            col_cnt += 1
    WRITE_CNT += 1




def mainn():
    #########################################################################################
    # 策略：
    #     type在每个pdf里面是唯一的，先搜索type，定位type在文件中的行号
    #     然后定位type在line中大概的位置
    #     然后在依次后面的行中，定位item的位置
    #     然后在命中item的行中，在item位置后，找到对应的数值
    #########################################################################################
    global WRITE_CNT
    global COL_CNT
    WRITE_CNT = 0
    FINAL_SETTING = SETTING
    xls_name = ''
    print u'请选择对应的数字:'
    print u'0. 输出2015之前的信息.'
    print u'1. 输出2015年 Nitrogen 相关的信息.'
    print u'2. 输出2015年 Phosphates 相关的信息.'
    print u'3. 输出2015年 Ammonia 相关的信息.'
    print u'4. 输出2015年 Sulphur 相关的信息.'
    print u'5. 输出2015年 Potash 相关的信息.'
    op = raw_input().lower()
    if op == '1':
        FINAL_SETTING = SETTING
        xls_name = '2014'
    elif op == '2':
        FINAL_SETTING = SETTING2015
        xls_name = '2015Nitrogen'
    elif op == '2':
        FINAL_SETTING = SETTING2015B
        xls_name = '2015Phosphates'
    elif op == '2':
        FINAL_SETTING = SETTING2015C
        xls_name = '2015Ammonia'
    elif op == '2':
        FINAL_SETTING = SETTING2015D
        xls_name = '2015Sulphur'
    elif op == '2':
        FINAL_SETTING = SETTING2015E
        xls_name = '2015Potash'
    else:
        FINAL_SETTING = SETTING
        xls_name = '2014'
    
    
    work_sheet = None
    # 创建一个xls文件
    cur_time = time.strftime('%Y%m%d-%H%M',time.localtime(time.time()))
    xls_name = xls_name + '_' + cur_time+'.xls'
    xls_file = xlwt.Workbook(encoding ='gbk')
    work_sheet = xls_file.add_sheet('1', cell_overwrite_ok=True)
    xls_init2015(work_sheet, FINAL_SETTING)
    xls_file.save(xls_name)

    # return
    # start
    for pdf in os.listdir(PDF_FILE_DIR):
        # pdf = '20150327fmbipg.pdf'
        print 'START:', pdf
        date = pdf[0:8]
        COL_CNT = 0
        # if '2015' not in pdf:
        # if '20130418' not in pdf:
        if FINAL_SETTING == SETTING:
            if '2015' in pdf:
                continue
        else:
            if '2015' not in pdf:
                continue

        pdf2txt(PDF_FILE_DIR + pdf)

        fn = open(DES_FILE, 'r')
        fn_lines = [i.lower() for i in fn.readlines()]
        for j in FINAL_SETTING[0:]:                                       # 调试位置1
            print '------START:', j['type']
            line_cnt = 0
            m = j['type'][0]
            # 找到匹配的行数
            for (index, cur_line) in enumerate(fn_lines):
                if match_detect(m, cur_line):                 # 寻找表头
                    line_cnt = index
                    print 'line_cnt:', line_cnt
                    break

            position = fn_lines[line_cnt].lower().find(m.lower().split(' ')[0])     # 在一行中的位置
            print position
            ret_list = info_extract(j['items'], fn_lines[line_cnt:], position)       # 调试位置2

            write2xls(work_sheet, ret_list, date)
            xls_file.save(xls_name)

            date = None
            # return
        if fn:
            fn.close()
        WRITE_CNT += 1
        # return
    if os.path.exists(DES_FILE):
        os.remove(DES_FILE)

def main():
    #########################################################################################
    # 策略：
    #     type在每个pdf里面是唯一的，先搜索type，定位type在文件中的行号
    #     然后定位type在line中大概的位置
    #     然后在依次后面的行中，定位item的位置
    #     然后在命中item的行中，在item位置后，找到对应的数值
    #########################################################################################
    global WRITE_CNT
    global COL_CNT
    WRITE_CNT = 0
    
    print u'请选择对应的数字:'
    print u'1. 输出2015之前的信息.'
    print u'2. 输出2015年'
    # 创建一个xls文件
    cur_time = time.strftime('%Y%m%d-%H%M',time.localtime(time.time()))
    xls_name = cur_time+'.xls'
    xls_file = xlwt.Workbook(encoding ='gbk')
    work_sheet = xls_file.add_sheet('1', cell_overwrite_ok=True)
    xls_init(work_sheet)
    # start
    for pdf in os.listdir(PDF_FILE_DIR):
        # pdf = '20140306Wk10FMBInternationalPriceGuide.pdf'

        date = pdf[0:8]
        COL_CNT = 0
        if '2015' in pdf:
            continue
        print 'START:', pdf
        
        pdf2txt(PDF_FILE_DIR + pdf)

        fn = open(DES_FILE, 'r')
        fn_lines = [i.lower() for i in fn.readlines()]
        for j in SETTING[0:]:                                       # 调试位置1
            print '------START:', j['type']
            line_cnt = 0
            m = j['type'][0]
            # 找到匹配的行数
            for (index, cur_line) in enumerate(fn_lines):
                if match_detect(m, cur_line):                 # 寻找表头
                    line_cnt = index
                    # print 'line_cnt:', line_cnt
                    break

            position = fn_lines[line_cnt].lower().find(m.lower().split(' ')[0])
            ret_list = info_extract(j['items'], fn_lines[line_cnt:], position)       # 调试位置2

            write2xls(work_sheet, ret_list, date)
            xls_file.save(xls_name)

            date = None
            # return
        if fn:
            fn.close()
        WRITE_CNT += 1
        # return


def test():
    ##########################################################################################
    # 测试直接查找，是否会存在重复的情况
    # 测试结果: item 存在重复的情况
    #           type 不存在重复的情况
    ##########################################################################################
    temp_list = [0, 0, 0]
    for pdf in os.listdir(PDF_FILE_DIR):
        if '2015' in pdf:
            continue
        pdf2txt(PDF_FILE_DIR + pdf)
        for j in SETTING:
            for m in j['type']:
                match_cnt = 0
                line_cnt = 0
                fn = open(DES_FILE, 'r')
                for i in fn.readlines():
                    line_cnt += 1
                    if match_detect(m, i):
                        match_cnt += 1
                if match_cnt == 0:
                    print 'PDF:', pdf
                    print 'NO Match.', 'type:', j['type'], '. item:', m
                    temp_list[0] += 1
                elif match_cnt == 1:
                    print 'PDF:', pdf
                    print 'Match Once.', 'type:', j['type'], '. item:', m
                    temp_list[1] += 1
                else:
                    print 'PDF:', pdf
                    print 'Too More.', 'type:', j['type'], '. item:', m
                    temp_list[2] += 1
                if fn:
                    fn.close()
        # return
    print temp_list


if __name__ == "__main__":
    # run()
    # test()
    # main()
    mainn()
    
    # info_extract()







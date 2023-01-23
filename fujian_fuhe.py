import re
# final = [wh, bihou, mode, have_gaiban, tibian, zhonglei,classes]

def not_empty(s):
    return s and s.strip()

def get_wh(val):#得到长宽
    wh = []
    for data in val:
        n = "".join(filter(lambda s: s in '0123456789', data))
        if len(n)==3 or n == '50' or n == '1000' or n == '1200':
            wh.append(n)
    wh = filter(not_empty, wh)
    wh = list(wh)
    wh = [wh[0],wh[1]]
    return wh

def baoku(vals):
    wh = get_wh(vals)
    wh = wh[0] + '×' + wh[1]
    bihou = vals[-1]
    bihou = float(bihou)
    zhonglei = 'baoku'
    final = [wh,bihou,'','','',zhonglei,'直通桥架']
    return final
def fengtou(vals):
    wh = get_wh(vals)
    wh = wh[0] + '×' + wh[1]
    bihou = vals[-1]
    bihou = float(bihou)
    zhonglei = 'fengtou'
    final = [wh,bihou,'','','',zhonglei,'直通桥架']
    return final
def guanjietou(vals):
    mode = 'buxiugang'
    wh = ''
    zhonglei = 'guanjietou'
    for data in vals:
        if data.find('铝合金'):
            mode = 'lvhejin'
        n = "".join(filter(lambda s: s in '0123456789', data))
        if n != '':
            wh = 'DN' + n
    final = [wh,'',mode,'','',zhonglei,'']
    return final
def zadai(vals):
    number = []
    for data in vals:
        n = "".join(filter(lambda s: s in '0123456789', data))
        if n != '':
            number.append(n)
    wh = number[0]+'×'+number[1]
    zhonglei = 'zadai'
    final = [wh,'','','','',zhonglei,'']
    return final
def geban(vals):
    wh = 0
    wh_list = []
    for data in vals:
        n = "".join(filter(lambda s: s in '0123456789', data))
        if n != '':
            wh_list.append(n)
    zhonglei  = 'geban'
    wh = wh_list[0]
    bihou = wh_list[1]
    final = [wh,bihou,'','','',zhonglei,'']
    return final
def zhijiepian(vals):
    wh_list = []
    for data in vals:
        n = "".join(filter(lambda s: s in '0123456789', data))
        wh_list.append(str(n))
    wh_list = filter(not_empty, wh_list)
    wh_list = list(wh_list)
    wh = wh_list[0]
    zhonglei = 'zhijiepian'
    final = [wh,'','','','',zhonglei,'']
    return final
def wanjiepian(vals):
    wh_list = []
    for data in vals:
        n = "".join(filter(lambda s: s in '0123456789', data))
        wh_list.append(str(n))
    wh_list = filter(not_empty, wh_list)
    wh_list = list(wh_list)
    wh = wh_list[0]
    zhonglei = 'wanjiepian'
    final = [wh, '', '', '', '', zhonglei,'']
    return final
def tiaojiaopian(vals):
    wh_list = []
    for data in vals:
        n = "".join(filter(lambda s: s in '0123456789', data))
        wh_list.append(str(n))
    wh_list = filter(not_empty, wh_list)
    wh_list = list(wh_list)
    wh = wh_list[0]
    zhonglei = 'tiaojiaopian'
    final = [wh, '', '', '', '', zhonglei,'']
    return final
def lianjiexian(vals):
    wh_list = []
    wh = 0
    for data in vals:
        n = "".join(filter(lambda s: s in '0123456789', data))
        if n != '':
            wh_list.append(n)
    if wh_list[0] != '1':
        wh = wh_list[0]
    else:
        wh = wh_list[1]
    zhonglei = 'lianjiexian'
    final = [wh,'','','','',zhonglei,'']
    return final
def gudingyaban(vals):
    zhonglei = 'gudingyaban'
    final = ['','','','','',zhonglei,'']
    return final
def xiangjiaodian(vals):
    zhonglei = 'xiangjiaodian'
    final = ['', '', '', '', '', zhonglei,'']
    return final
def shensuopian(vals):
    zhonglei = 'shensuopian'
    wh = 0
    wh_list = []
    for data in vals:
        n = "".join(filter(lambda s: s in '0123456789', data))
        if n != '':
            wh_list.append(n)
    wh = wh_list[0]
    final = [wh,'', '', '', '', zhonglei, '']
    return final
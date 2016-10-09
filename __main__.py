__author__ = 'weinaguo'
import openpyxl
import pandas as pd

headers = ['weight', 'rate', 'zone', 'from', 'to', 'commitment']

def envelopeupto8oz(wb):
    env = wb.get_sheet_by_name('EnvelopeUpTo8oz')
    firstovernight = [env['A'+str(x)].value for x in range(2, env.max_row+1)]     # List comprehension
    zone = [env['E'+str(x)].value for x in range(2, env.max_row+1)]
    ret = pd.DataFrame()
    ret['rate'] = firstovernight
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def envelopeupto8oz(wb):
    env = wb.get_sheet_by_name('EnvelopeUpTo8oz')
    priorityovernight = [env['B'+str(x)].value for x in range(2, env.max_row+1)]
    zone = [env['E'+str(x)].value for x in range(2, env.max_row+1)]
    ret2 = pd.DataFrame()
    ret2['rate'] = priorityovernight
    ret2['zone'] = zone
    for h in headers:
        if h not in ret2:
            ret2[h] = [0 for _ in range(len(ret2['rate']))]
    return ret2

def envelopeupto8oz(wb):
    env = wb.get_sheet_by_name('EnvelopeUpTo8oz')
    standardovernight = [env['C'+str(x)].value for x in range(2, env.max_row+1)]
    zone = [env['E'+str(x)].value for x in range(2, env.max_row+1)]
    ret3 = pd.DataFrame()
    ret3['rate'] = standardovernight
    ret3['zone'] = zone
    for h in headers:
        if h not in ret3:
            ret3[h] = [0 for _ in range(len(ret3['rate']))]
    return ret3

def envelopeupto8oz(wb):
    env = wb.get_sheet_by_name('EnvelopeUpTo8oz')
    twodayam = [env['D'+str(x)].value for x in range(2, env.max_row+1)]
    zone = [env['E'+str(x)].value for x in range(2, env.max_row+1)]
    ret4 = pd.DataFrame()
    ret4['rate'] = twodayam
    ret4['zone'] = zone
    for h in headers:
        if h not in ret4:
            ret4[h] = [0 for _ in range(len(ret4['rate']))]
    return ret4

def pfirstovernight(wb):
    pfo = wb.get_sheet_by_name('ExpressPackageFirstOvernight')
    weight = [pfo['A'+str(x)].value for x in range(2, pfo.max_row+1)]
    firstovernight = [pfo['B'+str(x)].value for x in range(2, pfo.max_row+1)]
    zone =[pfo['C'+str(x)].value for x in range(2, pfo.max_row+1)]
    ret5 = pd.DataFrame()
    ret5['weight'] = weight
    ret5['rate'] = firstovernight
    ret5['zong'] = zone
    for h in headers:
        if h not in ret5:
            ret5[h] = [0 for _ in range(len(ret5['rate']))]
    return ret5

def ppriorityovernight(wb):
    ppo = wb.get_sheet_by_name('ExpressPackagePriorityOvernight')
    weight = [ppo['A'+str(x)].value for x in range(2, ppo.max_row+1)]
    priorityovernight = [ppo['B'+str(x)].value for x in range(2, ppo.max_row+1)]
    zone =[ppo['C'+str(x)].value for x in range(2, ppo.max_row+1)]
    ret6 = pd.DataFrame()
    ret6['weight'] = weight
    ret6['rate'] = priorityovernight
    ret6['zone'] = zone
    for h in headers:
        if h not in ret6:
            ret6[h] = [0 for _ in range(len(ret6['rate']))]
    return ret6

def pstandardovernight(wb):
    pso = wb.get_sheet_by_name('ExpressPackageStandardOvernight')
    weight = [pso['A'+str(x)].value for x in range(2, pso.max_row+1)]
    standardovernight = [pso['B'+str(x)].value for x in range(2, pso.max_row+1)]
    zone = [pso['C'+str(x)].value for x in range(2, pso.max_row+1)]
    ret7 = pd.DataFrame()
    ret7['weight'] = weight
    ret7['rate'] = standardovernight
    ret7['zone'] = zone
    for h in headers:
        if h not in ret7:
            ret7[h] = [0 for _ in range(len(ret7['rate']))]
    return ret7

def ptwodayam(wb):
    p2a = wb.get_sheet_by_name('ExpressPackage2DayAM')
    weight = [p2a['A'+str(x)].value for x in range(2, p2a.max_row+1)]
    rate = [p2a['B'+str(x)].value for x in range(2, p2a.max_row+1)]
    zone = [p2a['C'+str(x)].value for x in range(2, p2a.max_row+1)]
    ret8 = pd.DataFrame()
    ret8['weight'] = weight
    ret8['rate'] = rate
    ret8['zone'] = zone
    for h in headers:
        if h not in ret8:
            ret8[h] = [0 for _ in range(len(ret8['rate']))]
    return ret8

def ptwoday(wb):
    p2 = wb.get_sheet_by_name('ExpressPackage2Day')
    weight = [p2['A'+str(x)].value for x in range(2, p2.max_row+1)]
    rate = [p2['B'+str(x)].value for x in range(2, p2.max_row+1)]
    zone = [p2['C'+str(x)].value for x in range(2, p2.max_row+1)]
    ret9 = pd.DataFrame()
    ret9['weight'] = weight
    ret9['rate'] = rate
    ret9['zone'] = zone
    for h in headers:
        if h not in ret9:
            ret9[h] = [0 for _ in range(len(ret9['rate']))]
    return ret9

def psaver(wb):
    ps = wb.get_sheet_by_name('ExpressPackageSaver')
    weight = [ps['A'+str(x)].value for x in range(2, ps.max_row+1)]
    rate = [ps['B'+str(x)].value for x in range(2, ps.max_row+1)]
    zone = [ps['C'+str(x)].value for x in range(2, ps.max_row+1)]
    ret10 = pd.DataFrame()
    ret10['weight'] = weight
    ret10['rate'] = rate
    ret10['zone'] = zone
    for h in headers:
        if h not in ret10:
            ret10[h] = [0 for _ in range(len(ret10['rate']))]
    return ret10

def groundandhomedelivery(wb):
    ghd = wb.get_sheet_by_name('GroundAndHomeDelivery')
    weight = [ghd['A'+str(x)].value for x in range(2, ghd.max_row+1)]
    rate = [ghd['B'+str(x)].value for x in range(2, ghd.max_row+1)]
    zone = [ghd['C'+str(x)].value for x in range(2, ghd.max_row+1)]
    commitment = [ghd['D'+str(x)].value for x in range(2,ghd.max_row+1)]
    ret11 = pd.DataFrame()
    ret11['weight'] = weight
    ret11['rate'] = rate
    ret11['zone'] = zone
    ret11['commitment'] = commitment
    for h in headers:
        if h not in ret11:
            ret11[h] = [0 for _ in range(len(ret11['rate']))]
    return ret11

def expressmultiweigt(wb):
    emw = wb.get_sheet_by_name('ExpressMultiweight')
    ws = [emw['A'+str(x)].value.split(u'\u2013')
          if u'\u2013' in emw['A'+str(x)].value else [emw['A'+str(x)].value[:-1], 0]
          for x in range(3, emw.max_row+1)]
    w_from, w_to = [x[0] for x in ws], [x[1] for x in ws]
    rate = [emw['B'+str(x)].value for x in range(3, emw.max_row+1)]
    zone = [emw['C'+str(x)].value for x in range(3, emw.max_row+1)]

    ret = pd.DataFrame()
    ret['from'] = w_from
    ret['to'] = w_to
    ret['rate'] = rate
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def sameday():
    ret = pd.DataFrame()
    weight = [x for x in range(70)]
    rate = [235 if x < 25 else 235+(x-24)*1.35 for x in weight]
    ret['weight'] = weight
    ret['rate'] = rate
    return ret

rules = {
    'envelopeupto8oz': envelopeupto8oz,
    'pfirstovernight': pfirstovernight,
    'ppriorityovernight': ppriorityovernight,
    'pstandardovernight': pstandardovernight,
    'ptwodayam': ptwodayam,
    'ptwoday': ptwoday,
    'psaver':psaver,
    'groundandhomedelivery':groundandhomedelivery

}



def main():
    wb = openpyxl.load_workbook('/Users/weinaguo/Desktop/2017fedexrates.xlsx')
    print envelopeupto8oz(wb)
    print pfirstovernight(wb)
    print ppriorityovernight(wb)
    print pstandardovernight(wb)
    print ptwodayam(wb)
    print ptwoday(wb)
    print psaver(wb)
    print groundandhomedelivery(wb)
    print expressmultiweigt(wb)

if __name__ == '__main__':
    main()
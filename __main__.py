__author__ = 'weinaguo'
import openpyxl
import pandas as pd

headers = ['weight', 'rate', 'zone', 'from', 'to', 'commitment']

def envelopeupto8oz(wb):
    env = wb.get_sheet_by_name('EnvelopeUpTo8oz')
    firstovernight = [env['A'+str(x)].value for x in range(2, env.max_row+1)]     # List comprehension
    zone = [env['B'+str(x)].value for x in range(2, env.max_row+1)]
    ret = pd.DataFrame()
    ret['rate'] = firstovernight
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def envelopeupto8ozp(wb):
    env = wb.get_sheet_by_name('EnvelopeUpTo8ozp')
    priorityovernight = [env['A'+str(x)].value for x in range(2, env.max_row+1)]
    zone = [env['B'+str(x)].value for x in range(2, env.max_row+1)]
    ret = pd.DataFrame()
    ret['rate'] = priorityovernight
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def envelopeupto8ozs(wb):
    env = wb.get_sheet_by_name('EnvelopeUpTo8ozs')
    standardovernight = [env['A'+str(x)].value for x in range(2, env.max_row+1)]
    zone = [env['B'+str(x)].value for x in range(2, env.max_row+1)]
    ret = pd.DataFrame()
    ret['rate'] = standardovernight
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def envelopeupto8oz2(wb):
    env = wb.get_sheet_by_name('EnvelopeUpTo8oz2')
    twodayam = [env['A'+str(x)].value for x in range(2, env.max_row+1)]
    zone = [env['B'+str(x)].value for x in range(2, env.max_row+1)]
    ret = pd.DataFrame()
    ret['rate'] = twodayam
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def pfirstovernight(wb):
    pfo = wb.get_sheet_by_name('ExpressPackageFirstOvernight')
    weight = [pfo['A'+str(x)].value for x in range(2, pfo.max_row+1)]
    firstovernight = [pfo['B'+str(x)].value for x in range(2, pfo.max_row+1)]
    zone =[pfo['C'+str(x)].value for x in range(2, pfo.max_row+1)]
    ret = pd.DataFrame()
    ret['weight'] = weight
    ret['rate'] = firstovernight
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def ppriorityovernight(wb):
    ppo = wb.get_sheet_by_name('ExpressPackagePriorityOvernight')
    weight = [ppo['A'+str(x)].value for x in range(2, ppo.max_row+1)]
    priorityovernight = [ppo['B'+str(x)].value for x in range(2, ppo.max_row+1)]
    zone = [ppo['C'+str(x)].value for x in range(2, ppo.max_row+1)]
    ret = pd.DataFrame()
    ret['weight'] = weight
    ret['rate'] = priorityovernight
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def pstandardovernight(wb):
    pso = wb.get_sheet_by_name('ExpressPackageStandardOvernight')
    weight = [pso['A'+str(x)].value for x in range(2, pso.max_row+1)]
    standardovernight = [pso['B'+str(x)].value for x in range(2, pso.max_row+1)]
    zone = [pso['C'+str(x)].value for x in range(2, pso.max_row+1)]
    ret = pd.DataFrame()
    ret['weight'] = weight
    ret['rate'] = standardovernight
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def ptwodayam(wb):
    p2a = wb.get_sheet_by_name('ExpressPackage2DayAM')
    weight = [p2a['A'+str(x)].value for x in range(2, p2a.max_row+1)]
    rate = [p2a['B'+str(x)].value for x in range(2, p2a.max_row+1)]
    zone = [p2a['C'+str(x)].value for x in range(2, p2a.max_row+1)]
    ret = pd.DataFrame()
    ret['weight'] = weight
    ret['rate'] = rate
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def ptwoday(wb):
    p2 = wb.get_sheet_by_name('ExpressPackage2Day')
    weight = [p2['A'+str(x)].value for x in range(2, p2.max_row+1)]
    rate = [p2['B'+str(x)].value for x in range(2, p2.max_row+1)]
    zone = [p2['C'+str(x)].value for x in range(2, p2.max_row+1)]
    ret = pd.DataFrame()
    ret['weight'] = weight
    ret['rate'] = rate
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def psaver(wb):
    ps = wb.get_sheet_by_name('ExpressPackageSaver')
    weight = [ps['A'+str(x)].value for x in range(2, ps.max_row+1)]
    rate = [ps['B'+str(x)].value for x in range(2, ps.max_row+1)]
    zone = [ps['C'+str(x)].value for x in range(2, ps.max_row+1)]
    ret = pd.DataFrame()
    ret['weight'] = weight
    ret['rate'] = rate
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def groundandhomedelivery(wb):
    ghd = wb.get_sheet_by_name('GroundAndHomeDelivery')
    weight = [ghd['A'+str(x)].value for x in range(2, ghd.max_row+1)]
    rate = [ghd['B'+str(x)].value for x in range(2, ghd.max_row+1)]
    zone = [ghd['C'+str(x)].value for x in range(2, ghd.max_row+1)]
    commitment = [ghd['D'+str(x)].value for x in range(2, ghd.max_row+1)]
    ret = pd.DataFrame()
    ret['weight'] = weight
    ret['rate'] = rate
    ret['zone'] = zone
    ret['commitment'] = commitment
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def expressmultiweight(wb):
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

def expressmultiweight2(wb):
    emw = wb.get_sheet_by_name('ExpressMultiweight2')
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

def expressmultiweight3(wb):
    emw = wb.get_sheet_by_name('ExpressMultiweight3')
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

def expressmultiweight4(wb):
    emw = wb.get_sheet_by_name('ExpressMultiweight4')
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

def expressmultiweight5(wb):
    emw = wb.get_sheet_by_name('ExpressMultiweight5')
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

def expressmultiweight6(wb):
    emw = wb.get_sheet_by_name('ExpressMultiweight6')

    ws = [emw['A'+str(x)].value.split(u'-')
          if u'-' in emw['A'+str(x)].value else[emw['A'+str(x)].value[:-1], 0]
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

def firstovernightfreight2(wb):
    fof = wb.get_sheet_by_name('FirstOvernightFreight2')
    ws = [fof['A'+str(x)].value.split(u'\u2013')
          if u'\u2013' in fof['A'+str(x)].value else [fof['A'+str(x)].value[:-1], 0]
          for x in range(2, fof.max_row+1)]
    w_from, w_to = [x[0] for x in ws], [x[1] for x in ws]
    rate = [fof['B'+str(x)].value for x in range(2, fof.max_row+1)]
    zone = [fof['C'+str(x)].value for x in range(2, fof.max_row+1)]

    ret = pd.DataFrame()
    ret['from'] = w_from
    ret['to'] = w_to
    ret['rate'] = rate
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def onedayfreight(wb):
    odf = wb.get_sheet_by_name('1DayFreight')
    ws = [odf['A'+str(x)].value.split(u'\u2013')
          if u'\u2013' in odf['A'+str(x)].value else [odf['A'+str(x)].value[:-1], 0]
          for x in range(2, odf.max_row+1)]
    w_from, w_to = [x[0] for x in ws], [x[1] for x in ws]
    rate = [odf['B'+str(x)].value for x in range(2, odf.max_row+1)]
    zone = [odf['C'+str(x)].value for x in range(2, odf.max_row+1)]

    ret = pd.DataFrame()
    ret['from'] = w_from
    ret['to'] = w_to
    ret['rate'] = rate
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def twodayfreight(wb):
    tdf = wb.get_sheet_by_name('2DayFreight')
    ws = [tdf['A'+str(x)].value.split(u'\u2013')
          if u'\u2013' in tdf['A'+str(x)].value else [tdf['A'+str(x)].value[:-1], 0]
          for x in range(2, tdf.max_row+1)]
    w_from, w_to = [x[0] for x in ws], [x[1] for x in ws]
    rate = [tdf['B'+str(x)].value for x in range(2, tdf.max_row+1)]
    zone = [tdf['C'+str(x)].value for x in range(2, tdf.max_row+1)]

    ret = pd.DataFrame()
    ret['from'] = w_from
    ret['to'] = w_to
    ret['rate'] = rate
    ret['zone'] = zone
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

def threedayfreight(wb):
    thf = wb.get_sheet_by_name('3DayFreight')
    ws = [thf['A'+str(x)].value.split(u'\u2013')
          if u'\u2013' in thf['A'+str(x)].value else [thf['A'+str(x)].value[:-1], 0]
          for x in range(2, thf.max_row+1)]
    w_from, w_to = [x[0] for x in ws], [x[1] for x in ws]
    rate = [thf['B'+str(x)].value for x in range(2, thf.max_row+1)]
    zone = [thf['C'+str(x)].value for x in range(2, thf.max_row+1)]

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
    for h in headers:
        if h not in ret:
            ret[h] = [0 for _ in range(len(ret['rate']))]
    return ret

rules = {
    'envelopeupto8oz': envelopeupto8oz,
    'pfirstovernight': pfirstovernight,
    'ppriorityovernight': ppriorityovernight,
    'pstandardovernight': pstandardovernight,
    'ptwodayam': ptwodayam,
    'ptwoday': ptwoday,
    'psaver': psaver,
    'groundandhomedelivery': groundandhomedelivery

}



def main():
    wb = openpyxl.load_workbook('/Users/weinaguo/Desktop/2017fedexrates.xlsx')
    for rule in rules:
        pass
    results =[ envelopeupto8oz(wb)
    , pfirstovernight(wb)
    , ppriorityovernight(wb)
    , pstandardovernight(wb)
    , ptwodayam(wb)
    , ptwoday(wb)
    , psaver(wb)
    , groundandhomedelivery(wb)
    , expressmultiweight(wb)
    , expressmultiweight2(wb)
    , expressmultiweight3(wb)
    , expressmultiweight4(wb)
    , expressmultiweight5(wb)
    , expressmultiweight6(wb)
    , firstovernightfreight2(wb)
    , onedayfreight(wb)
    , twodayfreight(wb)
    , threedayfreight(wb)
    , sameday() ]
    result = pd.DataFrame()
    for r in results:
        result = result.append(r)
    writer = pd.ExcelWriter('/Users/weinaguo/Desktop/fedex2017-out.xlsx')
    result.to_excel(writer)
    writer.save()

if __name__ == '__main__':
    main()

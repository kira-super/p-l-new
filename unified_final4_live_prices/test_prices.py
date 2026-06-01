import requests, warnings
warnings.filterwarnings('ignore')

tickers = [
    # UAE - retry with exact Yahoo format
    ('AEA000201011','ADCB.AE'),('AEA000801018','ADIB.AE'),
    ('AEA006101017','ADNOCDRILL.AE'),('AEA007501017','YAHSAT.AE'),
    ('AEE000401019','EAND.AE'),('AEE000701012','EITC.AE'),
    ('AEE01134E227','EMCONSL.AE'),('AEE01135A222','AMERICANA.AE'),
    ('AEE01195A234','ADNOCGAS.AE'),('AEE01268A239','ADNOCLS.AE'),
    ('AEE01354I230','INVCORP.AE'),('AEE01356D236','DURTAXI.AE'),
    ('AEE01388A243','ALEF.AE'),('AEE01456N241','NMDC.AE'),
    ('AEE01657D252','DUBAIRESID.AE'),('AEN000101016','FAB.AE'),
    ('AEF000901015','FERTIGLOBE.AE'),
    # LSE still-active
    ('IE00BD5B1Y92','BOCH.L'),('JE00BLKGSR75','IDH.L'),
    ('GB00BYSS4K11','GHG.L'),('GB00BMXNWH07','NE'),
    ('NL0015002X96','THEON.AT'),
    # HK retry
    ('CNE100007HC1','2366.HK'),('CNE100007JY1','6843.HK'),
    ('CNE100007LD1','2143.HK'),('KYG0705A1085','6679.HK'),
    ('KYG9361H1092','0198.HK'),('KYG9808A1058','9871.HK'),
    # Saudi retry
    ('SA0007879360','1316.SR'),('SA15DGH1VOH4','4096.SR'),
    ('SA15GHD4KS19','2189.SR'),('SA15H14I11H6','4056.SR'),
    ('SA15QGU1UNH6','2404.SR'),('SA164H113MH2','2103.SR'),
    ('SA16502H3N15','2108.SR'),('SA562GSHUOH7','4025.SR'),
    # Poland retry
    ('PLCCC0000016','CCC.WA'),('PLDGNST00012','DIAG.WA'),
    ('PLGRPRC00015','PRC.WA'),('US44853H1086','HUUUGE.WA'),
    # Qatar retry
    ('QA000A0M6MD5','KCBK.QA'),('QA000PK2KD10','MEEZA.QA'),
    ('QA000QLM0003','QLM.QA'),
    # Romania
    ('ROSIFEACNOR4','SIF5.RO'),
    # Greece retry
    ('GRS001003052','CRED.AT'),('GRS419003009','OPAP.AT'),
    ('GRS518003009','IPTO.AT'),
    # Norway retry
    ('BMG6904D1083','PRAX.OL'),('MHY1096C1093','CTAN.OL'),
    ('KYG236271055','SHELF.OL'),
    # Brazil alternatives
    ('BRCPLEACNOR8','CPLE3.SA'),('BRELETACNOR6','ELET6.SA'),
    ('BRELMDACNOR3','ELMD3.SA'),('BREMBRACNOR4','ERJ'),
    ('BRSRNAACNOR4','SRNA3.SA'),
    # Mexico alternatives
    ('MX01BA1D0003','BACHOCOB.MX'),('MXCFFI170008','FMTY14.MX'),
    ('MXCFTE0B0005','TERRA13.MX'),('VGG0896C1032','TBBB'),
    # Thailand alternatives
    ('THA504010002','T1.BK'),('THA848010015','TLIFE.BK'),
    ('THB056010010','IATG.BK'),
    # Singapore/Japan retry
    ('SGXE54866863','8IH.SI'),('JP3758110005','3328.T'),
    # Kuwait retry
    ('KW0EQ0200653','KIPCO.KW'),('KW0EQ0609648','AGAB.KW'),
    ('KW0EQ0610166','AEC.KW'),
    # US OTC alternatives
    ('US44853H1086','HUU.F'),  # Frankfurt
    ('US42207L1061','HHRU'),
    ('US83418T1088','CIANF'),
    ('US48666V2043','KMG.L'),  # LSE
    ('US6565674016','NTL.BA'), # Buenos Aires
    ('US71646J1097','PZE.BA'),
    # Philippines
    ('PHY0040P1094','ALLHOME.PS'),('PHY077751022','BDO.PS'),
    ('PHY0927M1046','BLOOM.PS'),('PHY1244L1009','CHP.PS'),
    ('PHY1249R1024','CNPF.PS'),('PHY1757W1054','CNVRG.PS'),
    ('PHY6028G1361','MBT.PS'),('PHY7318T1017','RRHI.PS'),
    ('PHY8321B1036','SGP.PS'),('PH0000058638','OGP.PS'),
    ('PH0000061442','MWC.PS'),
]

for isin, t in tickers:
    try:
        r = requests.get(f'https://query1.finance.yahoo.com/v8/finance/chart/{t}',
            headers={'User-Agent':'Mozilla/5.0'},
            params={'interval':'1d','range':'5d'}, verify=False, timeout=8)
        closes = r.json()['chart']['result'][0]['indicators']['quote'][0]['close']
        closes = [x for x in closes if x]
        print(f'{isin:16} {t:20} {closes[-1]:.4f}')
    except:
        print(f'{isin:16} {t:20} FAIL')
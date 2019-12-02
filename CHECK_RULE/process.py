from pandas import isnull, notnull, Series
import re


def CheckFormat(df, name):
    
    print('\t----检查格式错误')

    def f(x):

        values = x.iloc[6:]

        if x.name in (2, 12):
            res = list(map(lambda x: notnull(x) and len(str(x)) <= 30, values))

        elif x.name in (3, 4, 5, 6, 7, 8, 10, 13, 14, 15, 16, 20, 21, 22, 26):
            res = list(map(lambda x: notnull(x), values))

        elif x.name in (11, ):
            res = list(map(lambda x: notnull(x) and x == name, values))

        elif x.name in (17, 18):
            # res = list(map(lambda x: notnull(x) and bool(re.match('\d{4}/\d{2}/\d{2}$', str(x))), values))
            res = list(map(lambda x: notnull(x), values))

        elif x.name in (25, 59, 60, 61, 63, 69, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 98, 99, 
        100, 101, 106, 110, 114, 123, 124, 126, 130, 135, 136, 137, 139, 146, 147, 148, 149, 150, 151, 152, 154, 155, 156,
        157, 159, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 177, 180, 181, 182, 188, 189, 190, 191, 192, 193,
        194, 195, 198, 199, 202, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 225,
        226, 227, 232, 234, 235, 236, 237, 238, 239, 240, 241, 242, 243, 245, 246, 248, 249, 250):
            key_values = x.iloc[5].split(',')
            res = list(map(lambda x: notnull(x) and x in key_values, values))

        elif x.name in (97, 143, 144, 145, 158, 183, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 0<=int(x)<=2, values))

        elif x.name in (95, 140, 153):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 0<=int(x)<=3, values))

        elif x.name in (125, 127, 141):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 0<=int(x)<=4, values))

        elif x.name in (128, 131, 133, 134, 184):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 0<=int(x)<=5, values))

        elif x.name in (23, 129, 132, 138, 142, 244, 247):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 0<=int(x)<=10, values))

        elif x.name in (102, 103):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 0<=int(x)<=20, values))

        elif x.name in (233, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 0<=int(x)<=34, values))

        elif x.name in (111, 112):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 0<=int(x)<=40, values))

        elif x.name in (64, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 10<=int(x)<=30, values))

        elif x.name in (65, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 100<=int(x)<=500, values))

        elif x.name in (66, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+$', str(x))) and 20<=int(x)<=90, values))

        elif x.name in (24, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 0<=float(x)<=10, values))

        elif x.name in (118, 231, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 0<=float(x)<=20, values))

        elif x.name in (256, 257, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 0<=float(x)<=100, values))

        elif x.name in (252, 253, 254, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 0<=float(x)<=1000, values))

        elif x.name in (48, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 1000<=float(x)<=4000, values))

        elif x.name in (49, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 2000<=float(x)<=4000, values))

        elif x.name in (47, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 1000<=float(x)<=5000, values))

        elif x.name in (46, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 1000<=float(x)<=8000, values))

        elif x.name in (27, ):
            res = list(map(lambda x: notnull(x) and bool(re.match('\d+(.\d+)?$', str(x))) and 0<=float(x)<=10000000, values))


        else:
            res = [None] * len(values)

        return Series(res, index=values.index)

    return df.apply(f, axis=1)



def CheckLogic(df):

    print('\t----检查逻辑错误')

    def f(x):

        nonlocal df
        _df = df.iloc[:, 6:]
        values = x.iloc[6:]
        res = []
        if x.name in (30, ):
            for fuel,this,a,b,c,d in zip(_df.loc[26], values, _df.loc[31], _df.loc[32], _df.loc[33], _df.loc[34]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append((this=='S' and all([x in ('-','Opt.') for x in [a,b,c,d]])) or (this in ('-','Opt.')))

        elif x.name in (31, ):
            for fuel,this,a,b,c,d in zip(_df.loc[26], values, _df.loc[30], _df.loc[32], _df.loc[33], _df.loc[34]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append((this=='S' and all([x in ('-','Opt.') for x in [a,b,c,d]])) or (this in ('-','Opt.')))

        elif x.name in (32, ):
            for fuel,this,a,b,c,d in zip(_df.loc[26], values, _df.loc[30], _df.loc[31], _df.loc[33], _df.loc[34]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append((this=='S' and all([x in ('-','Opt.') for x in [a,b,c,d]])) or (this in ('-','Opt.')))

        elif x.name in (33, ):
            for fuel,this,a,b,c,d in zip(_df.loc[26], values, _df.loc[30], _df.loc[31], _df.loc[32], _df.loc[34]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append((this=='S' and all([x in ('-','Opt.') for x in [a,b,c,d]])) or (this in ('-','Opt.')))

        elif x.name in (34, ):
            for fuel,this,a,b,c,d in zip(_df.loc[26], values, _df.loc[30], _df.loc[31], _df.loc[32], _df.loc[33]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append((this=='S' and all([x in ('-','Opt.') for x in [a,b,c,d]])) or (this in ('-','Opt.')))

        elif x.name in (35, ):
            for fuel,this,a,b,c,d in zip(_df.loc[26], values, _df.loc[30], _df.loc[32], _df.loc[33], _df.loc[34]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append((this=='S' and [a,b,c,d].count('S')==1) or this in ('-','Opt.'))

        elif x.name in (36, ):            # 遗漏情况 CT6 PHEV_20161229(NM).xls
            for fuel,this,a,b,c,d,e in zip(_df.loc[26], values, _df.loc[33], _df.loc[30], _df.loc[31], _df.loc[32], _df.loc[34]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(
                        bool(re.match('\d+$', str(this))) 
                        and ((a=='S' and int(this)==0) or ([b,c,d,e].count('S')==1 and 0<=int(this)<=10)))

        elif x.name in (37, ):
            for fuel,this,a,b in zip(_df.loc[26], values, _df.loc[33], _df.loc[36]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(
                        bool(re.match('\d+$', str(this)))
                        and ((a=='S' and int(this)==0)
                             or (int(b)>6 and int(this)==int(b)-6)
                             or int(this)==0))

        elif x.name in (39, ):
            for fuel,this in zip(_df.loc[26], values):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(bool(re.match('\d+(.\d+)?$', str(this))) and 0<=float(this)<=1000)

        elif x.name in (40, ):
            for fuel,this in zip(_df.loc[26], values):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(bool(re.match('\d+(.\d+)?$', str(this))) and 1<=float(this)<=1000)

        elif x.name in (41, 43, 44, 50):
            for fuel,this in zip(_df.loc[26], values):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(this in ('S','-','Opt.'))

        elif x.name in (42, ):
            for fuel,this in zip(_df.loc[26], values):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(bool(re.match('\d+$', str(this))) and 0<=int(this)<=8)

        elif x.name in (45, ):
            for fuel,this in zip(_df.loc[26], values):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(bool(re.match('\d+(.\d+)?$', str(this))) and 0<=float(this)<=50)

        elif x.name in (51, ):
            for fuel,this,a,b,c in zip(_df.loc[26], values, _df.loc[50], _df.loc[52], _df.loc[53]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(
                        (this=='S' and a=='S' and all([x in ('-','Opt.') for x in [b,c]]))
                        or this in ('-','Opt.'))

        elif x.name in (52, ):
            for fuel,this,a,b,c in zip(_df.loc[26], values, _df.loc[50], _df.loc[51], _df.loc[53]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(
                        (this=='S' and a=='S' and all([x in ('-','Opt.') for x in [b,c]]))
                        or this in ('-','Opt.'))

        elif x.name in (53, ):
            for fuel,this,a,b,c in zip(_df.loc[26], values, _df.loc[50], _df.loc[51], _df.loc[52]):
                if fuel == 'BEV':
                    res.append(isnull(this))
                else:
                    res.append(
                        (this=='S' and a=='S' and all([x in ('-','Opt.') for x in [b,c]]))
                        or this in ('-','Opt.'))

        elif x.name in (54, ):
            for this,a in zip(values, _df.loc[55]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-','Opt.'))

        elif x.name in (55, ):
            for this,a in zip(values, _df.loc[54]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-','Opt.'))

        elif x.name in (56, ):
            for this,a in zip(values, _df.loc[57]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-','Opt.'))

        elif x.name in (57, ):
            for this,a in zip(values, _df.loc[56]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-','Opt.'))

        elif x.name in (58, ):
            for this,a in zip(values, _df.loc[56]):
                res.append(
                    (this=='S' and a=='S') or this in ('-','Opt.'))

        elif x.name in (67, ):
            for this,a in zip(values, _df.loc[68]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-','Opt.'))

        elif x.name in (68, ):
            for this,a in zip(values, _df.loc[67]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-','Opt.'))

        elif x.name in (70, ):
            for this,a,b,c in zip(values, _df.loc[71], _df.loc[72], _df.loc[74]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (71, ):
            for this,a,b,c in zip(values, _df.loc[70], _df.loc[72], _df.loc[74]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (72, ):
            for this,a,b,c in zip(values, _df.loc[70], _df.loc[71], _df.loc[74]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (74, ):
            for this,a,b,c in zip(values, _df.loc[70], _df.loc[71], _df.loc[72]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (73, ):                                                               # 存疑
            for this,a,b,c in zip(values, _df.loc[70], _df.loc[71], _df.loc[72]):
                res.append(
                    (this=='S' and [a,b,c].count('S')==1)
                    or (this=='Opt.' and [a,b,c].count('S')==1)
                    or (this=='Opt.' and [a,b,c].count('Opt.')==1)
                    or this=='-')

        elif x.name in (91, ):
            for body,this in zip(_df.loc[21], values):
                res.append(
                    (body!='SUV' and this=='-')
                    or this in ('S','-','Opt.'))

        elif x.name in (92, ):
            for this,a,b,c in zip(values, _df.loc[93], _df.loc[94], _df.loc[96]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (93, ):
            for this,a,b,c in zip(values, _df.loc[92], _df.loc[94], _df.loc[96]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (94, ):
            for this,a,b,c in zip(values, _df.loc[92], _df.loc[93], _df.loc[96]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (96, ):
            for this,a,b,c in zip(values, _df.loc[92], _df.loc[93], _df.loc[94]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (104, 105):
            for body,this in zip(_df.loc[21], values):
                res.append(
                    (notnull(this) and bool(re.match('\d+$', str(this))) and 0<=int(this)<=10)
                    and ((body!='MPV' and int(this)==0) or body=='MPV'))

        elif x.name in (108, ):
            for this,a in zip(values, _df.loc[109]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-','Opt.'))

        elif x.name in (109, ):
            for this,a in zip(values, _df.loc[108]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-','Opt.'))

        elif x.name in (113, ):
            for this,a,b in zip(values, _df.loc[111], _df.loc[112]):
                res.append(
                    (this=='S' and 0<int(a)<=40 and int(b)==0)
                    or (this=='S' and 0<int(a)<=40 and 0<int(b)<=40)
                    or (this=='S' and int(a)==0 and 0<int(b)<=40)
                    or this in ('-','Opt.'))

        elif x.name in (115, ):
            for this,a,b in zip(values, _df.loc[116], _df.loc[117]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b]]))
                    or this in ('-', 'Opt.'))

        elif x.name in (116, ):
            for this,a in zip(values, _df.loc[115]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-', 'Opt.'))

        elif x.name in (117, ):
            for this,a,b in zip(values, _df.loc[116], _df.loc[115]):
                res.append(
                    (this in ('S','Opt.') and a=='S' and b in ('-','Opt.'))
                    or (this=='Opt.' and a=='Opt.')
                    or this=='-')

        elif x.name in (119, ):
            for this,a,b,c in zip(values, _df.loc[120], _df.loc[121], _df.loc[122]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-', 'Opt.'))

        elif x.name in (120, ):
            for this,a,b,c in zip(values, _df.loc[119], _df.loc[121], _df.loc[122]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-', 'Opt.'))

        elif x.name in (121, ):
            for this,a,b in zip(values, _df.loc[119], _df.loc[120]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b]]))
                    or this in ('-', 'Opt.'))

        elif x.name in (122, ):
            for this,a,b,c in zip(values, _df.loc[121], _df.loc[119], _df.loc[120]):
                res.append(
                    (this in ('S','Opt.') and a=='S' and all([x in ('-','Opt.') for x in [b,c]]))
                    or (this=='Opt.' and a=='Opt.')
                    or this in ('-', 'Opt.'))

        elif x.name in (160, ):
            for this,a in zip(values, _df.loc[161]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-', 'Opt.'))

        elif x.name in (161, ):
            for this,a in zip(values, _df.loc[160]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-', 'Opt.'))

        elif x.name in (174, ):
            for this,a,b in zip(values, _df.loc[175], _df.loc[176]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b]]))
                    or this in ('-', 'Opt.'))

        elif x.name in (175, ):
            for this,a in zip(values, _df.loc[174]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-', 'Opt.'))

        elif x.name in (176, ):
            for this,a,b in zip(values, _df.loc[175], _df.loc[174]):
                res.append(
                    (this in ('S','Opt.') and a=='S' and b in ('-','Opt.'))
                    or (this=='Opt.' and a=='Opt.')
                    or this=='-')

        elif x.name in (178, ):
            for this,a,b,c in zip(values, _df.loc[176], _df.loc[175], _df.loc[167]):
                res.append(
                    (this=='S' and a==b==c=='S')
                    or (this=='Opt.' and a=='S' and b=='S' and c in ('S','Opt.'))
                    or (this=='Opt.' and a=='Opt.' and b in ('S','Opt.') and c in ('S','Opt.'))
                    or this=='-')

        elif x.name in (179, ):
            for this,a,b,c,d in zip(values, _df.loc[178], _df.loc[176], _df.loc[175], _df.loc[167]):
                res.append(
                    (this=='S' and a==b==c==d=='S')
                    or (this=='Opt.' and a==b==c==d=='S')
                    or (this=='Opt.' and a=='Opt.' and b=='S' and c=='S' and d in ('S','Opt.'))
                    or (this=='Opt.' and a=='Opt.' and b=='Opt.' and c in ('S','Opt.') and d in ('S','Opt.'))
                    or this=='-')

        elif x.name in (185, ):
            for this,a in zip(values, _df.loc[186]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-', 'Opt.'))

        elif x.name in (186, ):
            for this,a in zip(values, _df.loc[185]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-', 'Opt.'))

        elif x.name in (187, ):
            for this,a,b in zip(values, _df.loc[186], _df.loc[185]):
                res.append(
                    (this in ('S','Opt.') and a=='S' and b in ('-','Opt.'))
                    or (this=='Opt.' and a=='Opt.')
                    or this=='-')

        elif x.name in (196, ):
            for this,a,b,c in zip(values, _df.loc[187], _df.loc[186], _df.loc[185]):
                res.append(
                    (this=='S' and a==b=='S' and c!='S')
                    or (this=='Opt.' and a in ('S','Opt.') and b=='S' and c!='S')
                    or (this=='Opt.' and a==b=='Opt.')
                    or this=='-')

        elif x.name in (197, ):
            for this,a,b,c,d in zip(values, _df.loc[196], _df.loc[187], _df.loc[186], _df.loc[185]):
                res.append(
                    (this=='S' and a==b==c=='S' and d!='S')
                    or (this=='Opt.' and a in ('S','Opt.') and b=='S' and c=='S' and d!='S')
                    or (this=='Opt.' and a==b=='Opt.' and c=='S' and d!='S')
                    or (this=='Opt.' and a==b==c=='Opt.')
                    or this=='-')

        elif x.name in (200, ):
            for this,a in zip(values, _df.loc[201]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-', 'Opt.'))

        elif x.name in (201, ):
            for this,a in zip(values, _df.loc[200]):
                res.append(
                    (this=='S' and a in ('-','Opt.')) or this in ('-', 'Opt.'))

        elif x.name in (221, ):
            for this,a,b,c in zip(values, _df.loc[222], _df.loc[223], _df.loc[224]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (222, ):
            for this,a,b,c in zip(values, _df.loc[221], _df.loc[223], _df.loc[224]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (223, ):
            for this,a,b,c in zip(values, _df.loc[221], _df.loc[222], _df.loc[224]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (224, ):
            for this,a,b,c in zip(values, _df.loc[221], _df.loc[222], _df.loc[223]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b,c]]))
                    or this in ('-','Opt.'))

        elif x.name in (228, ):
            for this,a,b in zip(values, _df.loc[229], _df.loc[230]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b]]))
                    or this in ('-','Opt.'))

        elif x.name in (229, ):
            for this,a,b in zip(values, _df.loc[228], _df.loc[230]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b]]))
                    or this in ('-','Opt.'))

        elif x.name in (230, ):
            for this,a,b in zip(values, _df.loc[228], _df.loc[229]):
                res.append(
                    (this=='S' and all([x in ('-','Opt.') for x in [a,b]]))
                    or this in ('-','Opt.'))

        elif x.name in (255, 259, 261, 262, ):
            for fuel,this in zip(_df.loc[26], values):
                if fuel == 'PHEV':
                    res.append(isnull(this))
                else:
                    res.append(this in ('S','-','Opt.'))

        else:
            res = [True] * len(x.iloc[6:])

        return Series(res, index=x.iloc[6:].index)
    
    return df.apply(f, axis=1)



def RGB(red, green, blue):
    assert 0 <= red <=255
    assert 0 <= green <=255
    assert 0 <= blue <=255
    return red + (green << 8) + (blue << 16)

colorYellow, colorRed, colorNone = RGB(255, 255, 0), RGB(255, 0, 0), RGB(255, 255, 255)
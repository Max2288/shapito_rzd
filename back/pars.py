import pandas as pd


pars = {}
new = pd.read_excel('new.xlsx')
for i in new.itertuples():
    ind, prof, comp, indc, course, description = i
    pars[prof] = pars.get(prof, {})
    pars[prof][comp] = pars[prof].get(comp, {})
    pars[prof][comp][indc] = pars[prof][comp].get(\
        indc, [[course, description]])

# парсит в формате {профессия: {компетенция: {индикатор: [[курс, описание]...]}}}
